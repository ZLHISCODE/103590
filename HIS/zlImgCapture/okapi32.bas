Attribute VB_Name = "okapi32"
Option Explicit
'==============================================================================
'
'       filename        okapi32.bas
'       purpose         API for dynamic loading okapi32.dll (Ok series image cards)
'       language        Microsoft Visual Basic 5.0
'       author          H. Peng
'       date            2000.5.18
'-----------------------
'       modify by       H. Peng
'       purpose         Add Three Function
'       language        Microsoft Visual Basic 5.0
'       date            2001.3
'
'-----------------------
'
'       Copyright (C).  All Rights Reserved.
'
'
'==============================================================================


'---------------okapi32.bas---------------------------------
'
' ok api32 header file for user
'
'---------------------------------------------------------

'----contant defines----

'--defines of ok series image board identity
'Mono series
Public Const OK_M10 = 1010
Public Const OK_M10N = 1010
Public Const OK_M10M = 1013
Public Const OK_M10F = 1011
Public Const OK_M10L = 1014
Public Const OK_M10H = 1012
Public Const OK_M20 = 1020
Public Const OK_M2KC = 1021

Public Const OK_M20H = 1022
Public Const OK_M30 = 1030
Public Const OK_M40 = 1040
Public Const OK_M50 = 1050
Public Const OK_M60 = 1060
Public Const OK_M70 = 1070
Public Const OK_M80 = 1080
Public Const OK_M80K = 1081

'--new updated series
Public Const OK_M10A = 1210
Public Const OK_M10B = 1213                       'OK_M10M:1013
Public Const OK_M10C = 1214                       'OK_M10L/N:1014
Public Const OK_M10D = 1215
Public Const OK_M10K = 1218                       'OK_M80K


Public Const OK_M20A = 1222                       'OK_M20H:1022
Public Const OK_M20B = 1223                       '110M
Public Const OK_M20C = 1224                       '160M
Public Const OK_M20D = 1225                       '2050M

Public Const OK_M40A = 1240                       'OK_M40:1041
Public Const OK_M40B = 1243                       '110M
Public Const OK_M40C = 1244                       '160M
Public Const OK_M40D = 1245                       '205M

Public Const OK_M60A = 1260                       'OK_M60:1060
Public Const OK_M60B = 1263                       '110M
Public Const OK_M60C = 1264                       '160M
Public Const OK_M60D = 1265                       '205M

'64 bits series
Public Const OK_DM20B = 1223                      '140M
Public Const OK_DM20C = 1224                      '160M
Public Const OK_DM20D = 1225                      '205M

'PCI_X seriesC
Public Const OK_XM20_140 = 1323

'Color series
Public Const OK_C20 = 2020
Public Const OK_C20C = 2021
Public Const OK_C30 = 2030
Public Const OK_C32 = 2032
Public Const OK_C33 = 2033
Public Const OK_C30S = 2031
Public Const OK_C40 = 2040
Public Const OK_C50 = 2050
Public Const OK_C70 = 2070
Public Const OK_C80 = 2080
Public Const OK_C80M = 2081
Public Const OK_C82 = 2082

'RGB series
Public Const OK_RGB10 = 3010
Public Const OK_RGB20 = 3020
Public Const OK_RGB30 = 3030

'Monitor Control series
Public Const OK_MC10 = 4010
Public Const OK_MC16 = 4016
Public Const OK_MC20 = 4020
Public Const OK_MC30 = 4030

'--new updated series
Public Const OK_C20A = 2220
Public Const OK_C30A = 2230
Public Const OK_C40A = 2240
Public Const OK_C50A = 2250
Public Const OK_C60A = 2260

Public Const OK_RGB10A = 3210
Public Const OK_RGB10B = 3213

Public Const OK_RGB20A = 3220
Public Const OK_RGB20B = 3223
Public Const OK_RGB20C = 3224

Public Const OK_RGB30A = 3230
Public Const OK_RGB30B = 3233
Public Const OK_RGB30C = 3234

Public Const OK_MC10A = 4210
Public Const OK_MC12A = 4212
Public Const OK_MC16A = 4216


'---usb series
Public Const OK_USB20A = 5220

'---pc104+ series
Public Const OK_PC10 = 5210

'---cPCI series
Public Const OK_CPC16A = 5230


'--error code

Public Const ERR_NOERROR = 0                  'no error
Public Const ERR_NOTFOUNDBOARD = 1            'not found available ok board

Public Const ERR_NOTFOUNDVXDDRV = 2           'not found ok vxd driver
Public Const ERR_NOTALLOCATEDBUF = 3          'not pre-allocated buffer from host memory
Public Const ERR_BUFFERNOTENOUGH = 4          'available buffer not enough requirment
Public Const ERR_BEYONDFRAMEBUF = 5           'capture iamge size beyond buffer

Public Const ERR_NOTFOUNDDRIVER = 6           'no driver found
Public Const ERR_NOTCORRECTDRIVER = 7         'driver not correct

Public Const ERR_MEMORYNOTENOUGH = 8          'memory not enough
Public Const ERR_FUNNOTSUPPORT = 9            'the function not support
Public Const ERR_OPERATEFAILED = 10           'something wrong with the operation

Public Const ERR_HANDLEAPIERROR = 11          'the handle to okapi function wrong
Public Const ERR_DRVINITWRONG = 12            'something wrong with driver initialize

Public Const ERR_RECTVALUEWRONG = 13          'the rect set wrong
Public Const ERR_FORMNOTSUPPORT = 14          'the form set not support by the board

Public Const ERR_TARGETNOTSUPPORT = 15        'the target not support by this function

Public Const ERR_NOSPECIFIEDBOARD = 16        'not found specified board correctly sloted

'--format defines
Public Const FORM_RGB888 = 1
Public Const FORM_RGB565 = 2
Public Const FORM_RGB555 = 3
Public Const FORM_RGB8888 = 4
Public Const FORM_RGB332 = 5
Public Const FORM_RGBAAA = 18

Public Const FORM_YUV422 = 6
Public Const FORM_YUV411 = 7
Public Const FORM_YUV16 = 8
Public Const FORM_YUV12 = 9
Public Const FORM_YUV9 = 10
Public Const FORM_YUV8 = 11

Public Const FORM_GRAY888 = 12
Public Const FORM_GRAY8888 = 13
Public Const FORM_GRAY8 = 14
Public Const FORM_GRAY10 = 15
Public Const FORM_GRAY12 = 16
Public Const FORM_GRAY16 = 17



'--mask command
Public Const MASK_DISABALE = 0                'turn of mask
Public Const MASK_POSITIVE = 1                '0 win clients visible, 1 video visible
Public Const MASK_NEGATIVE = 2                '0 for video 1 for win client (graph)


'--tv system standard
Public Const TV_PALSTANDARD = 0               'PAL
Public Const TV_NTSCSTANDARD = 1              'NTSC
Public Const TV_NONSTANDARD = 2               'NON_STD
Public Const TV_SECAMSTANDARD = 3             'SECAM


Public Const TV_PALMAXWIDTH = 768
Public Const TV_PALMAXHEIGHT = 576

Public Const TV_NTSCMAXWIDTH = 640            '720
Public Const TV_NTSCMAXHEIGHT = 480


'-----defines lParam for get param
Public Const GETCURRPARAM = -1

'-----sub-function defines for wParam of SetVideoParam
        'wParam cab be one of the follow
Public Const VIDEO_RESETALL = 0                 'reset all to sys default
Public Const VIDEO_SOURCECHAN = 1
                                                ' lParam=0,1.. Comp.Video; 0x100,101...to Y/C(S-Video), 0x200,0x201 to RGB Chan.Input
Public Const VIDEO_BRIGHTNESS = 2               ' LOWORD is brightness, for RGB exHiWord is channel (0:red, 1:green, 2:blue)
Public Const VIDEO_CONTRAST = 3                 ' LOWORD is contrast, for RGB exHiWord is channel (0:red, 1:green, 2:blue)
Public Const VIDEO_COLORHUE = 4
Public Const VIDEO_SATURATION = 5
Public Const VIDEO_RGBFORMAT = 6                ' when return low word  is code high word is bitcount
Public Const VIDEO_TVSTANDARD = 7               ' 0 PAL, 1 NTSC, 2 Non-stadard
Public Const VIDEO_SIGNALTYPE = 8               ' LOWORD 0 non-interlaced, 1 interlaced
                                                ' exHiWord 0 no slot in field header, 1 yes
Public Const VIDEO_RECTSHIFT = 9                ' malelong (x,y)
Public Const VIDEO_SYNCSIGCHAN = 10             ' LOWORD 0:Red,1:Grn,2:Blue, 3:Sync;
                                                ' exHiWord is source 0,1,..for RGB input,
                                                ' 0x100,0x101,... for comp.video input
                                                ' ( in this case LOWORD has not mean more)
Public Const VIDEO_AUXMONCHANN = 11             ' monitor video source chann on aux monitor
Public Const VIDEO_AVAILRECTSIZE = 12           ' exMakeLong(horz,vert)
                                                ' horz available pixels per scan line and
                                                ' vert available lines per frame
Public Const VIDEO_FREQSEG = 13                 ' set horz video frequency range
                                                ' 0:Low(7.5~15MHz),  1:middle(15~30), 2:High(30~60)
Public Const VIDEO_LINEPERIOD = 14              ' line period (in 0.5 us) generated by board
Public Const VIDEO_FRAMELINES = 15              ' lines per frame  generated by board

Public Const VIDEO_MISCCONTROL = 16             ' miscellaneous control bits
                                                ' b0:-satur, b1:- contr for c20, c30
                                                ' b2:agc
Public Const VIDEO_ENABLEGRAPHS = 17            ' enable graph
Public Const VIDEO_GAINADJUST = 18              ' gain adjust

    
    
'-----sub-function defines for wParam of SetCaptureParam
        'wParam cab be one of the follow
Public Const CAPTURE_RESETALL = 0           'reset all to sys default
Public Const CAPTURE_INTERVAL = 1
Public Const CAPTURE_CLIPMODE = 2           'LOWORD: clip mode when video and dest rect not same size
                                            'exHiWord: if captrure odd and even field crosslly
Public Const CAPTURE_SCRRGBFORMAT = 3       'when return, loword=code, exHiWord=bits
Public Const CAPTURE_BUFRGBFORMAT = 4
Public Const CAPTURE_FRMRGBFORMAT = 5
Public Const CAPTURE_BUFBLOCKSIZE = 6       'lParam=MAKELONG(width,height)
                                            'if set it 0 (default), the rect set by user will be as block size
Public Const CAPTURE_HARDMIRROR = 7         'bit0 x, bit1 y;
Public Const CAPTURE_VIASHARPEN = 8         'sample via sharpen filter
Public Const CAPTURE_VIAKFILTER = 9         'sample via recursion filter
Public Const CAPTURE_SAMPLEFIELD = 10       '0 in field (non-interlaced), 1 in frame (interlaced), (0,1 are basic)
                                            '2 in field but keep expend row,3 in field but interlaced one frame
                                            '(2,4 can affect only sampllng field(frame) by field(frame) )
                                            'in 3 up-dn frame
Public Const CAPTURE_HORZPIXELS = 11        ' set max horz pixel per scan line
Public Const CAPTURE_VERTLINES = 12         ' set max vert lines per frame

Public Const CAPTURE_ARITHMODE = 13         'arithmatic mode
Public Const CAPTURE_TO8BITMODE = 14        'the mode of high (eg. 10 bits) converted to 8bit
                                            'exHiWord(lParam)=0: linear scale,
                                            'exHiWord(lParam)!=0:clip mode, LOWORD(lParam)=offset
Public Const CAPTURE_SEQCAPWAIT = 15        ' bit0 if waiting finished for functions of sequence capturing and playbacking
                                            'bit1 if waiting finished capture then call callback function

Public Const CAPTURE_MISCCONTROL = 16       'miscellaneous control bits
                                            'bit0: 1: take one by one |okCapturByBuffer,okGetSeqCapture by interrupt control
                                            'bit1: 1: take last one   |
                                            
Public Const CAPTURE_TRIGCAPDELAY = 17      'set delay capture by trigger
Public Const CAPTURE_TURNCHANNELS = 18      'turn channel when sequence capture

Public Const SAMPLE_INFIELD = 0             'in field (non-interlaced)
Public Const SAMPLE_INFRAME = 1             'in frame of interlaced fields
                                            'the above two (0,1) are basic
Public Const SAMPLE_FIELDEXP = 2            'in field but expend (keep expend row)
Public Const SAMPLE_UPDNFRAME = 3           'in frame of up-downed fields
Public Const SAMPLE_FIELDINTER = 4          'in field but interlaced to one frame

'-----defines lParam for CAPTURE_CLIPMODE
Public Const RECT_SCALERECT = 0
Public Const RECT_CLIPCENTER = 1
Public Const RECT_FIXLEFTTOP = 2
        'in condition video rect great than screen rect:
        'if RECT_SCALERECT video rect will be scaled to match screen rect if it can. else
        'video rect will be adjusted to match screen rect
        '(1: center, take center video rect  2: left-top fixed, take same size rect)


'-----sub-function defines for lParam of GetSignalParam
Public Const SIGNAL_VIDEOEXIST = 1          '0 video  absent, 1 exist
Public Const SIGNAL_VIDEOTYPE = 2           '0 field, 1 interlaced
Public Const SIGNAL_SCANLINES = 3           'scan lines per frame
Public Const SIGNAL_LINEFREQ = 4            'line frequency
Public Const SIGNAL_FIELDFREQ = 5           'frame frequency
Public Const SIGNAL_FRAMEFREQ = 6           'frame frequency
Public Const SIGNAL_EXTTRIGGER = 7          'extern trigger status
Public Const SIGNAL_FIELDID = 8             'Field ID 0 odd, 1 even
Public Const SIGNAL_VIDEOCOLOR = 9          'color(1) or B/W(0)


'-----sub-function defines for lEvent of WaitSignalEvent
Public Const EVENT_FIELDHEADER = 1          'field header
Public Const EVENT_FRAMEHEADER = 2          'frame header
Public Const EVENT_ODDFIELD = 3             'odd field come
Public Const EVENT_EVENFIELD = 4            'even field come
Public Const EVENT_EXTTRIGGER = 5           'extern trigger come,
                                            '(exHiWord(lEvent) is index)

'-----sub-function defines for lParam of PutSignalParam
Public Const PUTSIGNAL_TRIGGER = 1          'put trigger signal, 1 trigger


'-----sub-function defines for lParam of okSetConvertParam
Public Const CONVERT_RESETALL = 0           'reset all to sys default
Public Const CONVERT_FIELDEXTEND = 1        'field extend
Public Const CONVERT_PALETTE = 2            'set convert palette (just for 8 to 24 or 32)
                    'lParam=0: restore system default, >0: new palette pointer
Public Const CONVERT_HORZEXTEND = 3         'horzental extend (integer times)
Public Const CONVERT_HORZSTRETCH = 4        'horzental stretch (arbitrary number times)
Public Const CONVERT_MIRROR = 5             'x and y mirror
Public Const CONVERT_UPRIGHT = 6            'up to righ(=1)(rotate right 90 D) or left (=2) (rotate left 90 D

'field extend mode
Public Const FIELD_JUSTCOPY = 0         'just copy row by row
Public Const FIELD_COPYEXTEND = 1       'copy one row and expend one row (x2)
Public Const FIELD_INTERLEAVE = 2       'just copy odd(1.) rows (/2)
Public Const FIELD_INTEREXTEND = 3      'copy one odd row and expend one row
Public Const FIELD_COPYINTERPOL = 4     'copy one odd row and interpolate one row
Public Const FIELD_INTERINTERPOL = 5    'copy odd row and interpolate even row

Public Const FIELD_INTEREVEN = 6        'just copy even(2.) rows (/2)
Public Const FIELD_INTEREXTEVEN = 7     'copy one even row and expend one row
Public Const FIELD_JUSTCOPYODD = 8      'just copy odd rows to odd rows
Public Const FIELD_JUSTCOPYEVEN = 9     'just copy even rows to even rows
Public Const FIELD_ODDEVENCROSS = 10    'copy odd and even crossly


                                    'just for the case without bit converting

'-----sub-function defines for wParam of okBeginEncode and okBeginDecode
Public Const CODE_JPEG = 1

'-----defines for several target we can support
'typedef LPARAM  TARGET;

Public Const BUFFER = 1           'Buffer(physical) allocated from host memory
Public Const VIDEO = 0            'Video source input to the board
Public Const VSCREEN = -1          'Screen supported by VGA
Public Const FRAME = -2           'Frame buffer on the board
Public Const MONITOR = -3         'Monitor supported by (D/A) TV standard


Public Const SEQFILE = &H5153             'SQ
Public Const BMPFILE = &H4D42             'BM

Public Const BLKHEADER = &H4B42           'BK
Public Const BMPHEADER = &H4D42           'BM
Public Const BUFHEADER = &H4642           'BF


'-----------struct defines---------------

Public Const LF_FACESIZE = 32

Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


'--app user used struct
Type tpBoardType
    iBoardTypeCode As Integer  'ok board type code (e.g. 2030)
    szBoardName As String * 20 'board name (eg."OK_M20H")
    iBoardRankCode As Integer
End Type
Public BOARDTYPE As tpBoardType
Public LPBOARDTYPE As Long  '24 bytes

'image file block size
Type tpBlockSize
    iWidth As Integer      'width
    iHeight As Integer     'height
    iBitCount As Integer   'pixel bytes iBitCount
    iFormType As Integer   'rgb format type, need to fill when RGB565 or RGB 555
    lBlockStep As Long     'block stride (step to next image header)
                           'need to fill when treat multi block else set 0
End Type
Public BLOCKSIZE As tpBlockSize

'image block info
Type tpBlockInfo
    iType As Integer        '=BK or SQ, BM
    'struct _blocksize;
    iWidth As Integer       'width
    iHeight As Integer      'height
    iBitCount As Integer    'pixel bytes iBitCount
    iFormType As Integer    'rgb format type, need to fill when RGB565 or RGB 555
    lBlockStep As Integer   'block stride (step to next image header)
    iHiStep  As Integer     'exHiWord of block stride
    lTotal As Integer       'frame num
    iHiTotal As Integer     'exHiWord of total
    iInterval As Integer    'frame interval
    
    lpBits As Long ' image data pointer / file path name
    lpExtra As Long ' extra data (like as palette, mask) pointer
End Type
Public BLOCKINFO As tpBlockInfo
Public LPBLOCKINFO As Long
Public blkUse As tpBlockInfo

'sequence file info
Type tpSeqInfo 'file info for seq
    iType As Integer   '=SQ or BM
    'struct _blocksize;
    iWidth As Integer      'width
    iHeight As Integer     'height
    iBitCount As Integer   'pixel bytes iBitCount
    iFormType As Integer   'rgb format type, need to fill when RGB565 or RGB 555
    lBlockStep As Integer  'block stride (step to next image header)
    iHiStep As Integer     'exHiWord of block stride

    lTotal As Integer      'frame num
    iHiTotal As Integer    'exHiWord of total
    iInterval As Integer   'frame interval
End Type
Public SEQINFO As tpSeqInfo

'for replay
Type tpDibInfo              'file info for seq
    lpbi As Long            'bitmap info
    lpdib As Long           'dib data
    hwndPlayBox As Long     '1 replaying, 0 quit
    iCurrFrame As Integer   'current frame in buffer
    iReserved As Integer
End Type
Public LPDIBINFO As Long
Public lpDibPlay As Long

'---image size------
Type tpImageSize
    dwWidth As Long
    dwHeight As Long
    dwBitCount As Long
    dwReserved1 As Long
End Type
Public IMAGESIZE As tpImageSize
Public LPIMAGESIZE As Long

'--jpeg set params
Type tpJpegParam
    dwSize As Long      ' the size of this strcut
    lpstrName As Long   'must be NULL if not use
    dwQuality As Long
    dwReserved1 As Long
End Type
Public JPEGPARAM As tpJpegParam
Public LPJPEGPARAM As Long

'---set text mode----
Type tpTextMode
    dwForeColor As Long     ' forecolor, see macro RGB in win
    dwBackColor As Long     ' backcolor, see macro RGB in win
    dwSetMode As Long       ' 0:FULLCOPY, 1: FULLXOR
    wFrameNo As Integer     ' place which frame of target
    wReserved As Integer    ' no used
End Type
Public SETTEXTMODE As tpTextMode


Public Const FULLCOPY = 0                 'copy full text region into target
Public Const FULLXOR = 1                  'xor full text region and target
Public Const COPYFONT = 2                 'just copy fonts strokes to target
Public Const XORFONT = 3                  'just xor fonts strokes and target


'------ okapi32 functions list -----------


'--1. basic routines--------------

'prolog and epilog
Declare Function okOpenBoard Lib "okapi32.dll" (iIndex As Long) As Long   'okLockBoard
        'open a Ok series board in specified index(0 based), return 0 if not found any
        'if success, return a handle to control specified board
        'if set index=-1, mean takes default index no. (default is 0
        'if user not specified by 'Ok Device Manager' in Control Pannel)
        'this index can be also a specified board type code
        'this function will change iIndex to the true used index,
        'if index input is -1 or type code

Declare Function okCloseBoard Lib "okapi32.dll" (ByVal hBoard As Long) As Boolean  'okCloseBoard
        'Unlock and close Ok board specified handle
Declare Function okGetLastError Lib "okapi32.dll" () As Long
        'Get last error msg

Declare Function okGetBufferSize Lib "okapi32.dll" (ByVal hBoard As Long, lpLinear As Long, dwSize As Long) As Long
        'get base address and size of pre-allocated buffer,
        'if success return the max. frame num in which can be store according to current set
        'else return false;
Declare Sub okGetBufferAddr Lib "okapi32.dll" (ByVal hApi As Long, ByVal iNoFrame As Long)
        'get base address of specified frame No. in BUFFER
        'if success return the linear base address
        'else return false;

Declare Function okGetTargetInfo Lib "okapi32.dll" (ByVal hApi As Long, ByVal tgt As Long, ByVal iNoFrame As Integer, wid As Integer, ht As Integer, stride As Long) As Long
        'get target info include base address, width, height and stride specified frame No.
        'if success return the linear base address and other infos, else return false;

Declare Function okGetTypeCode Lib "okapi32.dll" (ByVal hBoard As Long, lpBoardName As Byte) As Integer
        'return type code and name of specified handle


'set rect and capture
Declare Function okSetTargetRect Lib "okapi32.dll" (ByVal hBoard As Long, ByVal target As Long, lpTgtRect As RECT) As Long
        'set target (VIDEO, SCREEN, BUFFER, FRAME)capture to or from
        'if Rect.right or .bottom) are -1 , they will be filled current value
        'special note for target=BUFFER:
        'if never set CAPTURE_BUFBLOCKSIZE, the block size(W,H) of buffer will be changed
        'according to size of right x bottom of lpRect, else the size will not changed
        'if success return max frames this target can support, else return <=0

Declare Function okSetToWndRect Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hwnd As Long) As Boolean
        'set client rect of hwnd as screen rect


Declare Function okCaptureSingle Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal lStart As Long) As Boolean
        'capture video source to target which can be BUFFER, SCREEN, FRAME, MONITOR
        'start(o based).if success return 1, if failed return 0, if not support target -1
        'when this function sent command to grabber, then return immediately not wait to finish.
        'this function same as okCaptureTo(hBoard, Dest, wParam, 1);

Declare Function okCaptureActive Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal lStart As Long) As Boolean
        'capture continuous active video to same position in target which can be BUFFER, SCREEN, FRAME, MONITOR
        'start(o based).if success return 1, if failed return 0, if not support target -1
        'when this function sent continuous command to grabber, then return immediately
        'but note that some card like RGB30. when target is SCREEN, this function is a thread.
        'this function same as okCaptureTo(hBoard, Dest, wParam, 0);

Declare Function okCaptureThread Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal lStart As Long, ByVal lParam As Long) As Boolean
        'capture sequencely video to target which can be BUFFER, SCREEN, FRAME, MONITOR
        'start(o based). lParam>0: number of frame to capture to,
        'if lParam > total frames in BUFER, it will loop in rewind mode(i%total)
        'if lParam=-N mean it loop in buffer of N frame infinitely until call okStopCapture,
        'when -1 mean loop in all buffer.
        'return max num frame can be stored in the target if success,
        'return 0 if failed(eg. format not matched). -1 not support target
        'this call will create a thread to manage to capture sequencely then
        'return immediately not wait to finish. This thread will callback if need
        'this function same as okCaptureTo(hBoard, Dest, wParam, n);
        'but it is not same, when n=1 this function is also a thread and still support callback
        
Declare Function okCaptureToScreen Lib "okapi32.dll" (ByVal hBoard As Long) As Boolean
        'Start to capture to screen (video real-time display on screen) and return immediately
        'this is just a special routine of okCaptureTo(hBoard,SCREEN,0,0)

Declare Function okCaptureTo Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal start As Long, ByVal lParam As Long) As Boolean
        'capture video source to target which can be BUFFER, SCREEN, FRAME, MONITOR
        'start(o based), lParam>0: number of frame to capture to, =0: cont. mode,
        'if lParam > total frames in BUFER, it will loop in rewind mode(i%total)
        'if lParam=-1 mean it loop infinitely until call okStopCapture
        'return max num frame can be stored in the target if success,
        'return 0 if failed(eg. format not matched). -1 not support target
        'this call will return immediately not wait to finish.
        'This function is not recomended to use for new user
        '
Declare Function okPlaybackFrom Lib "okapi32.dll" (ByVal hBoard As Long, ByVal src As Long, ByVal start As Long, ByVal lParam As Long) As Boolean
        'playback on monitor from target which can be BUFFER, FRAME
        'start(o based), lParam>0: number of frame to capture to, =0: cont. mode
        'if lParam > total frames in BUFER, it will loop in rewind mode (i%total)
        'if lParam=-1 mean it loop infinitely until call okStopCapture
        'return max num frame be stored in the target if success,
        'return 0 if wrong. -1 not support target
        'this call will return immediately not wait to finish.
        '

'get status and stop capture
Declare Function okGetCaptureStatus Lib "okapi32.dll" (ByVal hBoard As Long, ByVal bWait As Boolean) As Long
        'query capturing status, if bWait then wait to finish capturing, else return immediately.
        'return 0 if finished, if cont. mode capturing return target capture to
        '(which include SCREEN -1, FRAME -2, MONITOR -3)
        'if capturing to/from BUFFER or file, return the frame No.(1 based) being capturing

Declare Function okStopCapture Lib "okapi32.dll" (ByVal hBoard As Long) As Boolean
        'Stop capturing to or playback from SCREEN, BUFFER or other targets
        'return target just captured to or from.
        'if capturing to/from BUFFER or file, return the frame No.(1 based) being capturing


Declare Function okGetSeqCapture Lib "okapi32.dll" (ByVal hApi As Long, ByVal start As Long, ByVal count As Long) As Long
        'get current frame no. of sequence capturing to buffer
        'start: set buffer no. to use, effecting only count==0
        'count: count no. to catpure, start from count=0
        
'capture by to /from
Declare Function okCaptureByBuffer Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal start As Integer, ByVal num As Long) As Boolean
        'capture sequence images to dest by way of two frame buffers (in BUFFER),
        'the frame size and format is taken as same as current config of BUFFER
        'if dest is file name which can be .seq or .bmp (will generate multi bmp files)
        'dest can be also a user memory pointer or a BLOCKINFO pointer (with user memory pointer)
        'retrun true immediately if success. num should be great than 0

        
Declare Function okCaptureByBufferEx Lib "okapi32.dll" (ByVal hBoard As Long, ByVal fileset As Long, ByVal dest As Long, ByVal start As Long, ByVal num As Long) As Boolean
        ' all are same as okCaptureByBuffer except for fileset which is quality when to jpg file

Declare Function okPlaybackByBuffer Lib "okapi32.dll" (ByVal hBoard As Long, ByVal src As Long, ByVal start As Long, ByVal num As Long) As Boolean
        'playback sequence images on monitor from src by way of two frame buffers (in BUFFER)
        'the size and format of BUFFER will be changed to same as src
        'src can be a file name which may be .seq or .bmp (first orderd bmp files)
        ''src can be also a BLOCKINFO pointer (with infos of user memory pointer,size and format)
        'if src is just user memory pointer, this function will think its block size and format
        'are same as current config of BUFFER (in this case can not support loop function).
        'retrun true immediately if success
        'if num is great than the true frame number in src, it will loop back
        'if num=-1 mean it will loop infinitely until call okStopCapture

'set and get params
Declare Function okSetVideoParam Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
        '----set video param sub-function defines
        'set video param and return previous param;
        'if input lParam=-1, just return previous param
        'if not support return -1, if error return -2

Declare Function okSetCaptureParam Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
        'set capture param and return previous param;
        'if input lParam=-1, just return previous param
        'if not support return -1, if error return -2


'transfer and convert rect
Declare Function okReadPixel Lib "okapi32.dll" (ByVal hBoard As Long, ByVal src As Long, ByVal start As Long, ByVal x As Integer, ByVal y As Integer) As Long
        'read value of one pixel specified (x,y) in frame start of src (SCREEN, BUFFER, FRAME...)
        'return is this pixel value, it may be with bits 8,16,24,or 32 depend on the src's format

Declare Function okWritePixel Lib "okapi32.dll" (ByVal hBoard As Long, ByVal tgt As Long, ByVal start As Long, ByVal x As Long, ByVal y As Long, ByVal lValue As Long) As Long
        'write value into specified (x,y) in the frame start of tgt (SCREEN, BUFFER, FRAME...)
        '
        
Declare Function okSetConvertParam Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
        'set convert param for for function okConvertRect
        'if not support return -1, if error return -2

Declare Function okReadRect Lib "okapi32.dll" (ByVal hApi As Long, ByVal src As Long, ByVal start As Long, lpbuf As Byte) As Long
        'read data into lpBuf from rect(set previous) in frame start of dst (SCREEN, BUFFER, FRAME)
        'the data in lpBuf stored in way row by row
        'if src not supported return -1, if failed return 0,
        'return -1 if not support, return 0 failed,
        'if success return data length read in byte
        'if lpBuf=NULL, just return data length to read

Declare Function okWriteRect Lib "okapi32.dll" (ByVal hApi As Long, ByVal dst As Long, ByVal start As Long, lpbuf As Byte) As Long
        'write data in lpBuf to rect(set previous) of dst (SCREEN, BUFFER, FRAME)
        'the data in lpBuf stored in way row by row
        'return -1 if not support, return 0 failed,
        'if success return data length written in byte

Declare Function okReadRectEx Lib "okapi32.dll" (ByVal hApi As Long, ByVal src As Long, ByVal start As Long, lpbuf As Byte, lParam As Long) As Long
        'read data into lpBuf from rect(set previous) in frame start of dst (SCREEN, BUFFER, FRAME)
        'the data in lpBuf stored in way row by row
        'if src not supported return -1, if failed return 0,
        'return -1 if not support, return 0 failed,
        'if success return data length read in byte
        'if lpBuf=NULL, just return data length to read
        'LOWORD(lParam）is form code for bits of lpBuf (e.g.：FORM_GRAY8），if it is 0 mean: as same as src
        'exHiWord(LParam) is the mode of taking channels. mode=0 take all, =1 red, =2 green, =3 blue;

Declare Function okWriteRectEx Lib "okapi32.dll" (ByVal hApi As Long, ByVal dst As Long, ByVal start As Long, lpbuf As Byte, lParam As Long) As Long
        'write data in lpBuf to rect(set previous) of dst (SCREEN, BUFFER, FRAME)
        'the data in lpBuf stored in way row by row
        'return -1 if not support, return 0 failed,
        'if success return data length written in byte
        'LOWORD(lParam）is form code for bits of lpBuf (e.g.：FORM_GRAY8），if it is 0 mean: as same as dst
        'exHiWord(LParam) is the mode of taking channels. mode=0 take all, =1 red, =2 green, =3 blue;

Declare Function okTransferRect Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dest As Long, ByVal iFirst As Long, ByVal src As Long, ByVal iStart As Long, ByVal lNum As Long) As Long
        'transfer source rect to dest rect (here target can be SCREEN, BUFFER, FRAME, also BLOCKINFO point to user memory)
        'if total in dest or src less than lNum, it will rewind to begin then continue
        'this function transfer in format of src, that means it don't convert pixel bits if dst and src are not same
        'if src or dst not supported return -1, if failed return 0,
        'if success return data length of one block image in byte

Declare Function okConvertRect Lib "okapi32.dll" (ByVal hApi As Long, ByVal dst As Long, ByVal first As Long, ByVal src As Long, ByVal start As Long, ByVal no As Long) As Long
        'transfer source rect to dest rect (here target can be SCREEN, BUFFER, FRAME, also BLOCKINFO point to user memory)
        'if total in dest or src < lNum, it will rewind to begin then continue
        'this function convert to pixel foramt of dst if dst has not same bits format as src
        'if src or dst not supported return -1, if failed return 0,
        'if success return image size of one block in pixel

Declare Function okConvertRectEx Lib "okapi32.dll" (ByVal hDstBoard As Long, ByVal dst As Long, ByVal first As Long, ByVal hSrcBoard As Long, ByVal src As Long, ByVal start As Long, ByVal no As Long) As Long
        'same as the above function okConvertRect except with src handle

Declare Function okSetTextTo Lib "okapi32.dll" (ByVal hBoard As Long, ByVal target As Long, lpRect As RECT, lfLogFont As LOGFONT, textmode As tpTextMode, ByVal lpString As String, ByVal lLength As Long) As Boolean
        'set text into the image of specified target
        'target can be SCREEN, BUFFER, FRAME, also BLOCKINFO
        'lpRect: just use its (left,top) to points start posision
        'lfLogFont: windows font definition, see window's document
        'textMode: specify forecolor, backcolor and set mode
        'lpString, lLength: text string and it's length

    
Declare Function okDrawEllipsTo Lib "okapi32.dll" (ByVal hBoard As Long, ByVal target As Long, ByVal lStart As Long, lpRect As RECT, ByVal iForeColor As Long) As Long
        'draw a ellips into specified target
        'target can be SCREEN, BUFFER, FRAME, also BLOCKINFO
        'lpRect: specify the rect region of expected ellips
        'iForeColor: draw value on to target
        'return the pixel count of ellips

Declare Function okCreateDCBitmap Lib "okapi32.dll" (ByVal hBoard As Long, ByVal target As Long, hDCBitmap As Long) As Long
        'create a memory DC compatible with windows's GDI, draw graphic and text etc. with GDI functions
        'target can be SCREEN, BUFFER, FRAME, also BLOCKINFO
        'hDCBitmap: return a handle with which we can map the graphics on the memory DC
        'to our target.
        'return the memory DC with which you can use windows"s GDI functions

Declare Function okMapDCBitmapTo Lib "okapi32.dll" (ByVal hDCBitmap As Long, ByVal lStart As Long) As Long
        'map the graphics of memory DC created by okCreateDCBitmap into specified target
        'hDCBitmap: the handle created by okCreateDCBitmap

Declare Function okFreeDCBitmap Lib "okapi32.dll" (ByVal hDCBitmap As Long) As Boolean
        'free the allocated resource by okCreateDCBitmap
        'hDCBitmap: the handle created by okCreateDCBitmap



'get and put signals
Declare Function okGetSignalParam Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Long) As Long
        'Get specified param of video signal source
        'if not support return -1, if error return -2, else return param

Declare Function okWaitSignalEvent Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Long, ByVal lMilliSecond As Long) As Long
        'Wait specified signal come
        'lMilliSecond is time-out time in milliseconds for to wait
        'if lMilliSecond is zero, the function returns current state immediately
        'if lMilliSecond is INFINITE(-1) wait forever until event come
        'return -1 not support, 0 speicfied signal not come, 1 come


Declare Function okPutSignalParam Lib "okapi32.dll" (ByVal hBoard As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
        'put specified signal param
        'if not support return -1, if error return -2,
        
'treat callback functions
Declare Function okSetSeqProcWnd Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hwndMain As Long) As Boolean
        'set proc hwnd for receive message about sequence capture

Declare Function okSetSeqCallback Lib "okapi32.dll" (ByVal hBoard As Long, ByVal BeginProc As Long, ByVal SeqProc As Long, ByVal EndProc As Long) As Boolean
        'set callback function for multi-frame capturing function
        '(which are okCaptureTo, okCaptureFrom,okCaptureToFile, okCaptureFromFile)
        'see follow

''BOOL    CALLBACK BeginProc(HANDLE hBoard); 'user defined callback function
        'callback this function before to capture
''BOOL    CALLBACK SeqProc(HANDLE hBoard, long No); 'user defined callback function
        ' callback this function after finish capturing one frame
        ' No is the number(0 based) frame just finished or being playbacked.
''BOOL    CALLBACK EndProc(HANDLE hBoard); 'user defined callback function
        ' callback this function after end capturing

'save and load files
Declare Function okSaveImageFile Lib "okapi32.dll" (ByVal hBoard As Long, ByVal szFileName As String, ByVal first As Long, ByVal target As Long, ByVal start As Long, ByVal num As Long) As Long
    'here target can be BUFFER, SCREEN, FRAME or user buffer pointor
    '1.if ext name=".bmp":
    'create new file and than save one frame in start position of target as bmp file
    '
    '2.if ext name=".seq":
    'save no frame from (start) in target into (first) frame pos in seq(sequence) file in sequencely.
    'if the file already exist the function will not delete it, that mean old contents in the file will be kept.
    'So if you want create a new seq file with a existed file name you must delete before this call .
    '

Declare Function okLoadImageFile Lib "okapi32.dll" (ByVal hBoard As Long, ByVal szFileName As String, ByVal first As Long, ByVal target As Long, ByVal start As Long, ByVal num As Long) As Long
    'here target can be BUFFER, SCREEN, FRAME or user buffer pointor
    '1.if ext name=".bmp":
    'load one frame into start position of target from bmp file
    '
    '2.if ext name=".seq":
    'load no frame into (start) in target from (first) frame pos in seq(sequence) file in sequencely.



'apps setup dialog
Declare Function okOpenSetParamDlg Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hParentWnd As Long) As Boolean
    'dialog to setup video param
Declare Function okOpenSeqCaptureDlg Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hParentWnd As Long) As Boolean
    'dialog to capture sequence image

Declare Function okOpenReplayDlg Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hwnd As Long, ByVal src As Any, ByVal total As Long) As Long
    'open modeless dialog to replay sequence images(in BUFFER, USERBUF or seq file) on SCREEN or MONITOR
    
'--2. special routines supported by some cards--------------

'---overlay mask:
Declare Function okEnableMask Lib "okapi32.dll" (ByVal hBoard As Long, ByVal bMask As Long) As Long
        'bMask因为要输入0,1,2故不能声明为Boolean
        '0: disable mask; 1: positive mask, 2: negative mask
        'positive: 0 for win clients visible, 1 video visible
        'negative: 0 for video visible,  1 for win client (graph) visible
        'if bMask=-1 actually not set just get status previous set
        'return last mask status,
 
Declare Function okSetMaskRect Lib "okapi32.dll" (ByVal hBoard As Long, lpRect As RECT, ByVal lpMask As Long) As Long
        'Set mask rect(lpRect is relative to lpDstRect in SetScreenRect or
        'SetBufferRect, lpMask is mask code (in byte 0 or 1). one byte for one pixel
        'if lpMask==1, set all rect region in lpRect video visible
        'if lpMask==0, set all rect region in lpRect video unvisible
        'return base linear address of inner mask bit

'---set out LUT:
Declare Function okFillOutLUT Lib "okapi32.dll" (ByVal hBoard As Long, bLUT As Byte, ByVal start As Integer, ByVal num As Integer) As Long
        'fill specified playback out LUT.
        'bLut stored values to fill, (r0,g0,b0, r1,g1,b1 ...)
        'start: offset pos in LUT(based 0), num: num items to fill


'set input LUT:
Declare Function okFillInputLUT Lib "okapi32.dll" (ByVal hBoard As Long, bLUT As Byte, ByVal start As Integer, ByVal num As Integer) As Long
        'fill specified input LUT.
        'bLut stored values to fill, (r0,g0,b0, r1,g1,b1 ...)
        'start: offset pos in LUT(based 0), num: num items to fill
        
Declare Function okCaptureSequence Lib "okapi32.dll" (ByVal hBoard As Long, ByVal lStart As Long, ByVal lNoFrame As Long) As Boolean
        'capture sequence to BUFFER by way of Interrupt Service Routine
        'Note: Only M10 series, M20H,M40, M60, M30, M70 and RGB20 support this way
        'wParam=start(o based). lParam>0: number of frame to capture to,
        'if lParam > total frames in BUFER, it will loop in rewind mode(i%total)
        'when -1 mean loop in all buffer. infinitely until call okStopCapture,
        'return max num frame can be stored in the target in this way if success,
        'return 0 if failed(eg. format not matched). -1 not support target
        'this call will start a Interrupt Service Routine to manage to capture sequencely then
        'return immediately not wait to finish. This routine not support callback

Declare Function okPlaybackSequence Lib "okapi32.dll" (ByVal hBoard As Long, ByVal lStart As Long, ByVal lNoFrame As Long) As Boolean
        'playback on monitor from BUFFER
        'start(0 based), lNoFrame>0: number of frame to capture to,
        'if lParam > total frames in BUFER, it will loop in rewind mode (i%total)
        'if lParam=-1 mean it loop infinitely until call okStopCapture
        'return max num frame be stored in the target if success,
        'return 0 if wrong. -1 not support
        'this call will start a Interrupt Service Routine to manage to playback sequencely then
        'return immediately not wait to finish. This routine not support callback
        '
        
        
'---multi cards access--------
Declare Function okGetSlotBoard Lib "okapi32.dll" (lpBoardInfo As Long) As Integer
        'Query all Ok boards available in PCI bus, return total number
Declare Function okGetBoardIndex Lib "okapi32.dll" (ByVal szBoardName As String, ByVal iNo As Integer) As Integer
        'Get index (start 0) of specified board name string (it can also be typcode string)
        ' and order in same name (start 0),
        'return -1 if no this specified ok board
Declare Function okGetBoardName Lib "okapi32.dll" (ByVal lIndex As Long, ByVal szBoardName As String) As Integer
        'get the board code and name of the specified index
        'return the type code if success else return 0 if no card
        
'multi cards capture
Declare Function okMulCaptureTo Lib "okapi32.dll" (lphBaord As Long, ByVal dest As Long, ByVal start As Long, ByVal lParam As Integer) As Boolean
        'control multi boards to capture to target simultaneously, lphBaord are pointer of hBoard of multi board
        'other functions are same as okCaptureByBuffer

Declare Function okMulCaptureByBuffer Lib "okapi32.dll" (lphBaord As Long, ByVal tgt As Long, ByVal start As Long, ByVal num As Long) As Boolean
        'control multi boards to capture by buffer simultaneously, lphBaord are pointer of hBoard of multi board
        'other functions are same as okCaptureByBuffer

'multi channels:
Declare Function okLoadInitParam Lib "okapi32.dll" (ByVal hApi As Long, ByVal iChannNo As Integer) As Boolean
        'load specified chann (and as current chann.)of initial params
Declare Function okSaveInitParam Lib "okapi32.dll" (ByVal hApi As Long, ByVal iChannNo As Integer) As Boolean
        'save current init param to specified chann (and as current chann.)

'get and lock buffer
Declare Function okGetAvailBuffer Lib "okapi32.dll" (lpLinear As Long, dwSize As Long) As Long
        'Get free meomery buffers pre-allocated .
        'call it when user hope to access buffer directly or lock for one board
Declare Function okLockBuffer Lib "okapi32.dll" (ByVal hBoard As Long, ByVal dwSizeByte As Long, lpBasLinear As Long) As Long
        'Lock speicfiled size meomery buffers, then other handle can not use
Declare Function okUnlockAllBuffer Lib "okapi32.dll" () As Boolean
        'Unlock all buffer

'Mem Block Buffer appended to BUFFER
Declare Function okApplyMemBlock Lib "okapi32.dll" (ByVal dwBlockSize As Long, ByVal dwBlockNo As Long) As Long
        'apply mem block used as buffer appended to BUFFER
        'return the number of blocks allocated actually

Declare Function okFreeMemBlock Lib "okapi32.dll" () As Boolean
        'release appended MemBlock by okApplyMemBlock

Declare Function okGetMemBlock Lib "okapi32.dll" (ByVal hApi As Long, dwEachSize As Long, dwBlockNo As Long) As Long
        'get the number of MemBlock and size per block applied by okApplyMemBlock
        'and return the number can be as buffer as to cureent set size of BUFFER
        
Declare Function okLockMemBlock Lib "okapi32.dll" (ByVal hBoard As Long, ByVal lBlockNo As Long) As Long
        'lock number of MemBlock to specified handle

Declare Function okUnlockMemBlock Lib "okapi32.dll" () As Boolean
        'unlcok all locked MemBlock
        
'--4. apps utilities-----------------

'-- set pre-allocate buffer size in k byte
Declare Function okSetAllocBuffer Lib "okapi32.dll" (ByVal dwSize As Long) As Long
        'set the new size to preallocate in k bytes,
        'if new size is not same as current,
        'then the functuion will restart the window system

Declare Function okSetStaticVxD Lib "okapi32.dll" (ByVal lMode As Long) As Integer
        'lMode=0: check if static vxd registered.
        '=1: create static vxd register
        '=2: delete static vxd register

Declare Function okSetNTDriver Lib "okapi32.dll" (ByVal bCmd As Integer) As Integer
        'bCmd=0: check if nt driver installed.
        '=1: install nt driver
        '=2: remove nt driver

Declare Function okUnRegister Lib "okapi32.dll" (ByVal dwCmd As Long) As Boolean
        'uninstall all registered and generated infos

Declare Function okGetProgramInfo Lib "okapi32.dll" (ByVal iItem As Integer, lpString As Byte, ByVal iSize As Integer) As Long
        'get program info
        
Public Const PROGRAM = 1
Public Const VERSION = 2
Public Const PREFIX = 3
Public Const COMPANY = 4
Public Const TELFAX = 5
Public Const WEBEMAIL = 6

Declare Function okSetLangResource Lib "okapi32.dll" (ByVal langcode As Long) As Long      '1252 for English, 936 for Simple Chinese

'encode and decode
Declare Function okBeginEncode Lib "okapi32.dll" (ByVal hApi As Long, ByVal wCodeWay As Integer, ByVal lParam As Long) As Long
        'start to encode images. wCodeWay is JPEG or other compress,
        'lParam is address of parameters preset to encode
        'return a handle of encoder if sucessful, else  return 0
Declare Function okEncodeImage Lib "okapi32.dll" (ByVal hCoder As Long, ByVal src As Long, ByVal start As Long, ByVal lpData As Long, ByVal maxlen As Long) As Long
        'encode one frame image . src is source like BUFFER, SCREEN, FRAME and BLOCK.
        'lpData to store coded data,  maxlen is maximum length of lpCodedData
        'return the length coded data
Declare Function okEndEncode Lib "okapi32.dll" (ByVal hCoder As Long) As Long
        'end encode and release resources of encoder

Declare Function okBeginDecode Lib "okapi32.dll" (ByVal hApi As Long, ByVal wCodeWay As Integer, ByVal lpData As Long, ByVal lpImageInfo As Long) As Long
        'start to decode images. wCodeWay is JPEG or other compress,
        'lpData is coded data with header infos, lpImageInfo will return size info, it can also be NULL
        'return a handle of decoder if sucessful, else  return 0
Declare Function okDecodeImage Lib "okapi32.dll" (ByVal hCoder As Long, ByVal lpData As Long, length As Long, ByVal target As Long, ByVal start As Long) As Long
        'decode coded data to image. lpData is coded data,
        'length is input length of coded data, it also output length of real used data
        'target is target decode to like BUFFER, SCREEN, FRAME,BLOCK or memory pointer
        'return TRUE if one image finished, else return 0
Declare Function okEndDecode Lib "okapi32.dll" (ByVal hCoder As Long) As Long
        'end decoder and release resources of decoder

'protect
Declare Function okReadProtCode Lib "okapi32.dll" (ByVal hApi As Long, ByVal iIndex As Integer) As Long

Declare Function okWriteProtCode Lib "okapi32.dll" (ByVal hApi As Long, ByVal iIndex As Integer, ByVal code As Long) As Long

'--5. audio section routines--------------

'-----defines wParam in okSetAudioParam
Public Const AUDIO_RESETALL = 0             'reset all to sys default
Public Const AUDIO_SAMPLEFRQ = 1            'Sample rate, in samples per second
Public Const AUDIO_SAMPLEBITS = 2           'Bits per sample
Public Const AUDIO_INVOLUME = 3             'Audio input gain control
Public Const AUDIO_CALLINTERVAL = 4         'callback after interval times


'prolog and epilog
Declare Function okOpenAudio Lib "okapi32.dll" (ByVal hBoard As Long, ByVal lParam As Long) As Long
        'open audio device owned by the video capture board
        'hBoard is handle of image board, lParam reserved argument must ne set to 0
        'return handle of the audio device

Declare Function okCloseAudio Lib "okapi32.dll" (ByVal hAudio As Long) As Boolean
        'close audio device

'capture and stop
Declare Function okCaptureAudio Lib "okapi32.dll" (ByVal hAudio As Long, ByVal target As Long, ByVal lpfnUserProc As Long, ByVal lParam As Long) As Long
        'start to capture audio data, target can be BUFFER(audio data buffer) or file name
        'lpfnUserProc is callback function pointer, it must be NULL if not using callback
        ' lParam reserved argument must be set to 0
        'return the maximum times in miliseconds the inner audio data memory can be stored for

Declare Function WriteAudioProc Lib "okapi32.dll" (ByVal hAudio As Long, ByVal lpAudBuf As Long, ByVal length As Long) As Long
        'this callback function must written as the above protype by user
        'call this function when there are enough audio data (the length in byte)
        'call this function when capture ended with argument length=0;

Declare Function okStopCaptureAudio Lib "okapi32.dll" (ByVal hAudio As Long) As Boolean
        'stop capturing audio data
        'return total length of read out by okReadAudioData

'set and get audio
Declare Function okSetAudioParam Lib "okapi32.dll" (ByVal hAudio As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
        'set the parameters to sample audio,  wParam see above defines
        'return the new set value if success
        'if not support return -1, if error return -2
        'if input lParam=-1, just return previous param

Declare Function okReadAudioData Lib "okapi32.dll" (ByVal hAudio As Long, ByVal lpAudioBuf As Long, ByVal lSize As Long) As Long
        'read audio captured from inner data buffer
        'lpAudioBuf is your memory address to store
        'lSize is data length (in byte) you expect to read
        'return the data length (in byte) truelly read

'--6. ports io utilities--------------
' for Non-PCI IO cards on WinNT/Win2000
' you must call this function before using follow port io functions
Declare Function okSetPortBase Lib "okapi32.dll" (ByVal wPortBase As Integer, ByVal iPortCount As Integer) As Boolean
        'preset ports to use by setting port base address and port count
        'to use some ports (default port=0x300, count=4) must preset by calling this one
        'and these ports can be used correctly only after system restarted
        
        'this function only used for OK PCI GPIO20
Declare Function okGetGPIOPort Lib "okapi32.dll" (ByVal Index As Integer, wPortBase As Integer) As Integer
    'get port base and count of ok PCI GPIO20
    'index is the nomber (0 based) of gpio cards,
    'wPortBase: port base,  return port count if success else return 0

'utilities just for WinNT/Win2000
Declare Function okOutputByte Lib "okapi32.dll" (ByVal wPort As Integer, ByVal data As Byte) As Boolean
    'output a byte at specified port

Declare Function okOutputShort Lib "okapi32.dll" (ByVal wPort As Integer, ByVal data As Integer) As Boolean
    'output a word at specified port

Declare Function okOutputLong Lib "okapi32.dll" (ByVal wPort As Integer, ByVal data As Long) As Boolean
    'output a dword at specified port

'----------- input data at port
Declare Function okInputByte Lib "okapi32.dll" (ByVal wPort As Integer) As Byte
    'input a byte at specified port

Declare Function okInputShort Lib "okapi32.dll" (ByVal wPort As Integer) As Integer
    'input a word at specified port

Declare Function okInputLong Lib "okapi32.dll" (ByVal wPort As Integer) As Long
    'input a dword at specified port

Declare Function okGetAddrForVB Lib "okapi32.dll" (Value As Any) As Long
        'return array address for VB
        
'-----------get time elapse
Declare Function okGetTickCount Lib "okapi32.dll" () As Long
        'same as GetTickCount but exactly than it on Win2k

Declare Sub okSleep Lib "okapi32.dll" (ByVal dwMill As Long)
        'same as Sleep but exactly than it on Win2k

'-------------------end--------------------------


