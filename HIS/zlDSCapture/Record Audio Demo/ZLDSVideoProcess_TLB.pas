unit ZLDSVideoProcess_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 2010/12/25 10:36:51 from Type Library described below.

// ************************************************************************  //
// Type Lib: E:\Demo\Capture Com(Delphi)\ZLDSVideoProcess.ocx (1)
// LIBID: {B1790453-7708-48C1-B5CC-75255FA4B066}
// LCID: 0
// Helpfile: 
// HelpString: ZLDSVideoProcess Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
// Errors:
//   Error creating palette bitmap of (TTMCIAudio) : Error reading control bitmap
//   Error creating palette bitmap of (TTMCIPlayer) : Error reading control bitmap
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  ZLDSVideoProcessMajorVersion = 1;
  ZLDSVideoProcessMinorVersion = 0;

  LIBID_ZLDSVideoProcess: TGUID = '{B1790453-7708-48C1-B5CC-75255FA4B066}';

  IID_IDSCapture: TGUID = '{8D57B5E5-4B36-4496-8EE8-54EF55B6C58C}';
  DIID_IDSCaptureEvents: TGUID = '{EC14A323-3D09-443B-A23E-FD86909CD935}';
  IID_IDSParameterEnum: TGUID = '{48F3F5B1-BBA8-49A7-94B4-03195A880EF6}';
  IID_IDSPlay: TGUID = '{36F986E0-6834-40BB-A444-18613D51FC10}';
  DIID_IDSPlayEvents: TGUID = '{932BAAE5-451C-47B3-BD8E-43DF7C4EF698}';
  CLASS_DSPlay: TGUID = '{BC410BFE-ED4B-4DFD-8506-2D6CB2BBF564}';
  CLASS_DSCapture: TGUID = '{137D6CFF-36DB-4AB2-BD2C-AC279626A8F3}';
  CLASS_DSCapParameterEnum: TGUID = '{82A61469-457C-4654-AC4D-87676EB914BB}';
  IID_ITMCIAudio: TGUID = '{15923E95-8673-41A8-904E-9E12EFEC925C}';
  DIID_ITMCIAudioEvents: TGUID = '{B4593540-4604-4E09-90AA-9AB097805AD7}';
  CLASS_TMCIAudio: TGUID = '{E9F8B5F5-84D2-47CA-B2C0-ABC8E9B840A4}';
  IID_ITMCIPlayer: TGUID = '{B3C684B1-40D8-46FF-9429-BC07BEC99877}';
  DIID_ITMCIPlayerEvents: TGUID = '{926B0355-E7BD-4277-85BF-FAA7DBF10133}';
  CLASS_TMCIPlayer: TGUID = '{F38977AD-8F88-4AE4-BD08-57584273CFF6}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TxActiveFormBorderStyle
type
  TxActiveFormBorderStyle = TOleEnum;
const
  afbNone = $00000000;
  afbSingle = $00000001;
  afbSunken = $00000002;
  afbRaised = $00000003;

// Constants for enum TxPrintScale
type
  TxPrintScale = TOleEnum;
const
  poNone = $00000000;
  poProportional = $00000001;
  poPrintToFit = $00000002;

// Constants for enum TxMouseButton
type
  TxMouseButton = TOleEnum;
const
  mbLeft = $00000000;
  mbRight = $00000001;
  mbMiddle = $00000002;

// Constants for enum TMouseType
type
  TMouseType = TOleEnum;
const
  mtLeft = $00000000;
  mtMid = $00000001;
  mtRight = $00000002;

// Constants for enum TShowModel
type
  TShowModel = TOleEnum;
const
  smNormal = $00000000;
  smFit = $00000001;
  smStretch = $00000002;
  smAutoFitCut = $00000003;
  smWindAutoFit = $00000004;

// Constants for enum TVideoState
type
  TVideoState = TOleEnum;
const
  vsStop = $00000000;
  vsPlay = $00000001;
  vsPause = $00000002;

// Constants for enum TCapParameterPostion
type
  TCapParameterPostion = TOleEnum;
const
  cppLeftTop = $00000000;
  cppTopCenter = $00000001;
  cppRightTop = $00000002;
  cppRightCenter = $00000003;
  cppRightBottom = $00000004;
  cppBottomCenter = $00000005;
  cppLeftBottom = $00000006;
  cppLeftCenter = $00000007;
  cppScreenCenter = $00000008;

// Constants for enum TVideoProperty
type
  TVideoProperty = TOleEnum;
const
  vpVideoFile = $00000000;
  vpMajorTypeName = $00000001;
  vpSubTypeName = $00000002;
  vpFormatTypeName = $00000003;
  vpTimeFormatName = $00000004;
  vpVideoColorDepth = $00000005;
  vpVideoWidth = $00000006;
  vpVideoHeight = $00000007;
  vpStreamCount = $00000008;
  vpFrameRate = $00000009;
  vpTimeLen = $0000000A;
  vpFrameLen = $0000000B;

// Constants for enum TSnatchWay
type
  TSnatchWay = TOleEnum;
const
  swVMR = $00000000;
  swDEVICE = $00000001;

// Constants for enum TQualityType
type
  TQualityType = TOleEnum;
const
  qtBrightness = $00000000;
  qtContrast = $00000001;
  qtHue = $00000002;
  qtSaturation = $00000003;
  qtGamma = $00000004;
  qtWhiteBlance = $00000005;

// Constants for enum THideCfgItem
type
  THideCfgItem = TOleEnum;
const
  hciVideoDisplay = $00000001;
  hciImageCapture = $00000002;
  hciAdvanceCfg = $00000004;
  hciVideoShowWay = $00000008;
  hciVideoSnatchWay = $00000010;
  hciVideoState = $00000020;
  hciCaptureDevice = $00000040;
  hciVideoQuality = $00000080;
  hciVideoEncoder = $00000100;

// Constants for enum TAnimateType
type
  TAnimateType = TOleEnum;
const
  atQiu = $00000000;
  atMidi = $00000001;
  atMiWu = $00000002;
  atStar = $00000003;
  atLogon = $00000004;

// Constants for enum TAxMCIAudioState
type
  TAxMCIAudioState = TOleEnum;
const
  adsRun = $00000000;
  adsStop = $00000001;
  adsPause = $00000002;

// Constants for enum TAxBPS
type
  TAxBPS = TOleEnum;
const
  bps8 = $00000000;
  bps16 = $00000001;

// Constants for enum TAxChannels
type
  TAxChannels = TOleEnum;
const
  acMono = $00000000;
  acStereo = $00000001;

// Constants for enum TAxBufferCount
type
  TAxBufferCount = TOleEnum;
const
  buf2 = $00000000;
  buf4 = $00000001;
  buf6 = $00000002;
  buf8 = $00000003;
  buf10 = $00000004;
  buf12 = $00000005;
  buf14 = $00000006;
  buf16 = $00000007;

// Constants for enum TMp3CompressRate
type
  TMp3CompressRate = TOleEnum;
const
  MCR32 = $00000020;
  MCR40 = $00000028;
  MCR48 = $00000030;
  MCR56 = $00000038;
  MCR64 = $00000040;
  MCR80 = $00000050;
  MCR96 = $00000060;
  MCR112 = $00000070;
  MCR128 = $00000080;
  MCR144 = $00000090;
  MCR160 = $000000A0;
  MCR192 = $000000C0;
  MCR224 = $000000E0;
  MCR256 = $00000100;
  MCR320 = $00000140;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDSCapture = interface;
  IDSCaptureDisp = dispinterface;
  IDSCaptureEvents = dispinterface;
  IDSParameterEnum = interface;
  IDSParameterEnumDisp = dispinterface;
  IDSPlay = interface;
  IDSPlayDisp = dispinterface;
  IDSPlayEvents = dispinterface;
  ITMCIAudio = interface;
  ITMCIAudioDisp = dispinterface;
  ITMCIAudioEvents = dispinterface;
  ITMCIPlayer = interface;
  ITMCIPlayerDisp = dispinterface;
  ITMCIPlayerEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  DSPlay = IDSPlay;
  DSCapture = IDSCapture;
  DSCapParameterEnum = IDSParameterEnum;
  TMCIAudio = ITMCIAudio;
  TMCIPlayer = ITMCIPlayer;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PPUserType1 = ^IFontDisp; {*}
  PUserType1 = ^TCaptureParameter; {*}

  TCaptureParameter = packed record
    CaptureDeviceName: WideString;
    VideoAnalog: WideString;
    colorDepth: SYSINT;
    videoSize: WideString;
    Brightness: SYSINT;
    Contrast: SYSINT;
    Hue: SYSINT;
    WhiteBlance: SYSINT;
    encoderName: WideString;
    LimitLength: SYSINT;
    leftRate: Double;
    topRate: Double;
    heightRate: Double;
    widthRate: Double;
    Gamma: SYSINT;
    Saturation: SYSINT;
    VideoShowModel: TShowModel;
    SnatchWay: TSnatchWay;
    InputCrossbar: Integer;
    OutputCrossbar: Integer;
    ParameterState: WordBool;
    IsApplyImageCut: WordBool;
    IsConvertGrayImg: WordBool;
    IsTimeLimit: WordBool;
    IsShowState: WordBool;
    IsAutoBrightness: WordBool;
    IsAutoContrast: WordBool;
    IsAutoHue: WordBool;
    IsAutoGamma: WordBool;
    IsAutoSaturation: WordBool;
    IsAutoWhiteBlance: WordBool;
  end;

  TVideoSize = packed record
    Width: Integer;
    Height: Integer;
  end;


// *********************************************************************//
// Interface: IDSCapture
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8D57B5E5-4B36-4496-8EE8-54EF55B6C58C}
// *********************************************************************//
  IDSCapture = interface(IDispatch)
    ['{8D57B5E5-4B36-4496-8EE8-54EF55B6C58C}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_AutoScroll: WordBool; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    function Get_Color: OLE_COLOR; safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    function Get_Font: IFontDisp; safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    function Get_KeyPreview: WordBool; safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    function Get_Scaled: WordBool; safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    function Get_Active: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    function Get_HelpFile: WideString; safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    function Get_ScreenSnap: WordBool; safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    function Get_SnapBuffer: Integer; safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    function ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString; safecall;
    function StartPreview: WideString; safecall;
    procedure FreeRes; safecall;
    function CaptureBmpImageToFile(const fileName: WideString): WideString; safecall;
    function StartCaptureVideo(const fileName: WideString): WideString; safecall;
    function StopCaptureVideo(out videoFile: WideString): WideString; safecall;
    function Get_IsStretch: WordBool; safecall;
    procedure Set_IsStretch(Value: WordBool); safecall;
    function Get_IsShowState: WordBool; safecall;
    procedure Set_IsShowState(Value: WordBool); safecall;
    function Get_IsFullScreen: WordBool; safecall;
    procedure Set_IsFullScreen(Value: WordBool); safecall;
    function Get_IsAdjustWindowSize: WordBool; safecall;
    procedure Set_IsAdjustWindowSize(Value: WordBool); safecall;
    function StopPreview: WideString; safecall;
    function ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString; safecall;
    function Get_IsFit: WordBool; safecall;
    procedure Set_IsFit(Value: WordBool); safecall;
    function ShowVideoCapturePinCfg(parentHandle: Integer): WideString; safecall;
    function ShowVfwVideoSourceCfg(parentHandle: Integer): WideString; safecall;
    function ShowVfwVideoFormatCfg(parentHandle: Integer): WideString; safecall;
    function ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString; safecall;
    function Get_PreviewState: WordBool; safecall;
    function Get_CaptureState: WordBool; safecall;
    function ReadParameterFromFile: WideString; safecall;
    function CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString; safecall;
    function Get_IsEscKeyQuitFullScreen: WordBool; safecall;
    procedure Set_IsEscKeyQuitFullScreen(Value: WordBool); safecall;
    function Get_IsDblClickQuitFullScreen: WordBool; safecall;
    procedure Set_IsDblClickQuitFullScreen(Value: WordBool); safecall;
    function Get_IsClickQuitFullScreen: WordBool; safecall;
    procedure Set_IsClickQuitFullScreen(Value: WordBool); safecall;
    function Get_CurWidth: Integer; safecall;
    procedure Set_CurWidth(Value: Integer); safecall;
    function Get_CurHeight: Integer; safecall;
    procedure Set_CurHeight(Value: Integer); safecall;
    function Get_CurVideoWidth: Integer; safecall;
    procedure Set_CurVideoWidth(Value: Integer); safecall;
    function Get_CurVideoHeight: Integer; safecall;
    procedure Set_CurVideoHeight(Value: Integer); safecall;
    function RefreshWindow: WideString; safecall;
    function Get_ShowModel: TShowModel; safecall;
    procedure Set_ShowModel(Value: TShowModel); safecall;
    function Get_CapParameterWindPos: TCapParameterPostion; safecall;
    procedure Set_CapParameterWindPos(Value: TCapParameterPostion); safecall;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString; safecall;
    function QuitFullScreen: WideString; safecall;
    function Get_SnatchWay: TSnatchWay; safecall;
    procedure Set_SnatchWay(Value: TSnatchWay); safecall;
    function UpdateVideoQuailty: WideString; safecall;
    function SaveParameterToFile: WideString; safecall;
    function GetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    function SetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    function RePreview: WideString; safecall;
    function ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString; safecall;
    function Get_ParameterCfgFileName: WideString; safecall;
    procedure Set_ParameterCfgFileName(const Value: WideString); safecall;
    function Get_HideCfgItem: Integer; safecall;
    procedure Set_HideCfgItem(Value: Integer); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function CaptureImgToClipBoard: WideString; safecall;
    function ShowVfwCompressCfg(parentHandle: Integer): WideString; safecall;
    function ShowVideoCrossbarCfg(parentHandle: Integer): WideString; safecall;
    function CaptureBmpImage: IPictureDisp; safecall;
    function CaptureJpgImage(compressRate: Integer): IPictureDisp; safecall;
    function GetRealVideoSize: TVideoSize; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property AutoScroll: WordBool read Get_AutoScroll write Set_AutoScroll;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property AxBorderStyle: TxActiveFormBorderStyle read Get_AxBorderStyle write Set_AxBorderStyle;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Color: OLE_COLOR read Get_Color write Set_Color;
    property Font: IFontDisp read Get_Font write Set_Font;
    property KeyPreview: WordBool read Get_KeyPreview write Set_KeyPreview;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property PrintScale: TxPrintScale read Get_PrintScale write Set_PrintScale;
    property Scaled: WordBool read Get_Scaled write Set_Scaled;
    property Active: WordBool read Get_Active;
    property DropTarget: WordBool read Get_DropTarget write Set_DropTarget;
    property HelpFile: WideString read Get_HelpFile write Set_HelpFile;
    property ScreenSnap: WordBool read Get_ScreenSnap write Set_ScreenSnap;
    property SnapBuffer: Integer read Get_SnapBuffer write Set_SnapBuffer;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property AlignDisabled: WordBool read Get_AlignDisabled;
    property VisibleDockClientCount: Integer read Get_VisibleDockClientCount;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property IsStretch: WordBool read Get_IsStretch write Set_IsStretch;
    property IsShowState: WordBool read Get_IsShowState write Set_IsShowState;
    property IsFullScreen: WordBool read Get_IsFullScreen write Set_IsFullScreen;
    property IsAdjustWindowSize: WordBool read Get_IsAdjustWindowSize write Set_IsAdjustWindowSize;
    property IsFit: WordBool read Get_IsFit write Set_IsFit;
    property PreviewState: WordBool read Get_PreviewState;
    property CaptureState: WordBool read Get_CaptureState;
    property IsEscKeyQuitFullScreen: WordBool read Get_IsEscKeyQuitFullScreen write Set_IsEscKeyQuitFullScreen;
    property IsDblClickQuitFullScreen: WordBool read Get_IsDblClickQuitFullScreen write Set_IsDblClickQuitFullScreen;
    property IsClickQuitFullScreen: WordBool read Get_IsClickQuitFullScreen write Set_IsClickQuitFullScreen;
    property CurWidth: Integer read Get_CurWidth write Set_CurWidth;
    property CurHeight: Integer read Get_CurHeight write Set_CurHeight;
    property CurVideoWidth: Integer read Get_CurVideoWidth write Set_CurVideoWidth;
    property CurVideoHeight: Integer read Get_CurVideoHeight write Set_CurVideoHeight;
    property ShowModel: TShowModel read Get_ShowModel write Set_ShowModel;
    property CapParameterWindPos: TCapParameterPostion read Get_CapParameterWindPos write Set_CapParameterWindPos;
    property SnatchWay: TSnatchWay read Get_SnatchWay write Set_SnatchWay;
    property ParameterCfgFileName: WideString read Get_ParameterCfgFileName write Set_ParameterCfgFileName;
    property HideCfgItem: Integer read Get_HideCfgItem write Set_HideCfgItem;
    property AppHandle: Integer read Get_AppHandle write Set_AppHandle;
  end;

// *********************************************************************//
// DispIntf:  IDSCaptureDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8D57B5E5-4B36-4496-8EE8-54EF55B6C58C}
// *********************************************************************//
  IDSCaptureDisp = dispinterface
    ['{8D57B5E5-4B36-4496-8EE8-54EF55B6C58C}']
    property Visible: WordBool dispid 201;
    property AutoScroll: WordBool dispid 202;
    property AutoSize: WordBool dispid 203;
    property AxBorderStyle: TxActiveFormBorderStyle dispid 204;
    property Caption: WideString dispid -518;
    property Color: OLE_COLOR dispid -501;
    property Font: IFontDisp dispid -512;
    property KeyPreview: WordBool dispid 205;
    property PixelsPerInch: Integer dispid 206;
    property PrintScale: TxPrintScale dispid 207;
    property Scaled: WordBool dispid 208;
    property Active: WordBool readonly dispid 209;
    property DropTarget: WordBool dispid 210;
    property HelpFile: WideString dispid 211;
    property ScreenSnap: WordBool dispid 212;
    property SnapBuffer: Integer dispid 213;
    property DoubleBuffered: WordBool dispid 214;
    property AlignDisabled: WordBool readonly dispid 215;
    property VisibleDockClientCount: Integer readonly dispid 216;
    property Enabled: WordBool dispid -514;
    function ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString; dispid 217;
    function StartPreview: WideString; dispid 218;
    procedure FreeRes; dispid 219;
    function CaptureBmpImageToFile(const fileName: WideString): WideString; dispid 220;
    function StartCaptureVideo(const fileName: WideString): WideString; dispid 221;
    function StopCaptureVideo(out videoFile: WideString): WideString; dispid 222;
    property IsStretch: WordBool dispid 223;
    property IsShowState: WordBool dispid 224;
    property IsFullScreen: WordBool dispid 225;
    property IsAdjustWindowSize: WordBool dispid 226;
    function StopPreview: WideString; dispid 227;
    function ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString; dispid 228;
    property IsFit: WordBool dispid 229;
    function ShowVideoCapturePinCfg(parentHandle: Integer): WideString; dispid 230;
    function ShowVfwVideoSourceCfg(parentHandle: Integer): WideString; dispid 231;
    function ShowVfwVideoFormatCfg(parentHandle: Integer): WideString; dispid 232;
    function ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString; dispid 233;
    property PreviewState: WordBool readonly dispid 234;
    property CaptureState: WordBool readonly dispid 235;
    function ReadParameterFromFile: WideString; dispid 236;
    function CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString; dispid 237;
    property IsEscKeyQuitFullScreen: WordBool dispid 238;
    property IsDblClickQuitFullScreen: WordBool dispid 239;
    property IsClickQuitFullScreen: WordBool dispid 240;
    property CurWidth: Integer dispid 241;
    property CurHeight: Integer dispid 242;
    property CurVideoWidth: Integer dispid 243;
    property CurVideoHeight: Integer dispid 244;
    function RefreshWindow: WideString; dispid 245;
    property ShowModel: TShowModel dispid 246;
    property CapParameterWindPos: TCapParameterPostion dispid 247;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString; dispid 248;
    function QuitFullScreen: WideString; dispid 249;
    property SnatchWay: TSnatchWay dispid 250;
    function UpdateVideoQuailty: WideString; dispid 252;
    function SaveParameterToFile: WideString; dispid 253;
    function GetCaptureParameter(var parameter: {??TCaptureParameter}OleVariant): WideString; dispid 251;
    function SetCaptureParameter(var parameter: {??TCaptureParameter}OleVariant): WideString; dispid 254;
    function RePreview: WideString; dispid 255;
    function ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString; dispid 256;
    property ParameterCfgFileName: WideString dispid 257;
    property HideCfgItem: Integer dispid 258;
    property AppHandle: Integer dispid 259;
    function CaptureImgToClipBoard: WideString; dispid 260;
    function ShowVfwCompressCfg(parentHandle: Integer): WideString; dispid 261;
    function ShowVideoCrossbarCfg(parentHandle: Integer): WideString; dispid 262;
    function CaptureBmpImage: IPictureDisp; dispid 263;
    function CaptureJpgImage(compressRate: Integer): IPictureDisp; dispid 264;
    function GetRealVideoSize: {??TVideoSize}OleVariant; dispid 265;
  end;

// *********************************************************************//
// DispIntf:  IDSCaptureEvents
// Flags:     (4096) Dispatchable
// GUID:      {EC14A323-3D09-443B-A23E-FD86909CD935}
// *********************************************************************//
  IDSCaptureEvents = dispinterface
    ['{EC14A323-3D09-443B-A23E-FD86909CD935}']
    procedure OnActivate; dispid 201;
    procedure OnClick; dispid 202;
    procedure OnCreate; dispid 203;
    procedure OnDblClick; dispid 204;
    procedure OnDestroy; dispid 205;
    procedure OnDeactivate; dispid 206;
    procedure OnKeyPress(var Key: Smallint); dispid 207;
    procedure OnPaint; dispid 208;
    procedure OnMouseDown(button: SYSINT; shift: SYSINT; x: SYSINT; y: SYSINT); dispid 209;
    procedure OnMouseMove(shift: SYSINT; x: SYSINT; y: SYSINT); dispid 210;
    procedure OnMouseUp(button: SYSINT; shift: SYSINT; x: SYSINT; y: SYSINT); dispid 211;
    procedure OnKeyDown(var Key: SYSINT; shift: SYSINT); dispid 212;
    procedure OnKeyUp(var Key: SYSINT; shift: SYSINT); dispid 213;
    procedure OnResize; dispid 214;
    procedure OnGotFocus; dispid 215;
    procedure OnLostFocus; dispid 216;
    procedure OnVideoSizeChange(videoWidth: SYSINT; videoHieght: SYSINT; windowWidth: SYSINT; 
                                windowHeight: SYSINT); dispid 217;
    procedure OnMouseWheel(shift: SYSINT; wheelDelta: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                           var handled: WordBool); dispid 218;
    procedure OnMouseWheelDown(shift: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                               var handled: WordBool); dispid 219;
    procedure OnMouseWheelUp(shift: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                             var handled: WordBool); dispid 220;
  end;

// *********************************************************************//
// Interface: IDSParameterEnum
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {48F3F5B1-BBA8-49A7-94B4-03195A880EF6}
// *********************************************************************//
  IDSParameterEnum = interface(IDispatch)
    ['{48F3F5B1-BBA8-49A7-94B4-03195A880EF6}']
    function GetDeviceCount(var deviceCount: SYSINT): WideString; safecall;
    function GetDeviceName(deviceIndex: SYSINT; var deviceName: WideString): WideString; safecall;
    function GetEncoderCount(var encoderCount: SYSINT): WideString; safecall;
    function GetEncoderName(encoderIndex: SYSINT; var encoderName: WideString): WideString; safecall;
    function GetVideoQualityMaxValue(const deviceName: WideString; qualityType: TQualityType; 
                                     var maxValue: SYSINT): WideString; safecall;
    function GetVideoAnalogCount(var analogCount: SYSINT): WideString; safecall;
    function GetVideoAnalogName(analogIndex: SYSINT; var analogName: WideString): WideString; safecall;
    function GetVideoSizeCount(var sizeCount: SYSINT): WideString; safecall;
    function GetVideoSizeName(sizeIndex: SYSINT; var sizeName: WideString): WideString; safecall;
    function GetVideoColorDepthCount(var colorDepthCount: SYSINT): WideString; safecall;
    function GetVideoColorDepth(colorDepthIndex: SYSINT; var colorDepth: SYSINT): WideString; safecall;
    function CheckIsVfwDevice(const deviceName: WideString): WordBool; safecall;
    function CheckIsSupportVmr: WordBool; safecall;
    function VideoSizeConvert(const videoSize: WideString): TVideoSize; safecall;
    function GetIsSupportQuailtiCfg(const deviceName: WideString): WordBool; safecall;
  end;

// *********************************************************************//
// DispIntf:  IDSParameterEnumDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {48F3F5B1-BBA8-49A7-94B4-03195A880EF6}
// *********************************************************************//
  IDSParameterEnumDisp = dispinterface
    ['{48F3F5B1-BBA8-49A7-94B4-03195A880EF6}']
    function GetDeviceCount(var deviceCount: SYSINT): WideString; dispid 201;
    function GetDeviceName(deviceIndex: SYSINT; var deviceName: WideString): WideString; dispid 202;
    function GetEncoderCount(var encoderCount: SYSINT): WideString; dispid 203;
    function GetEncoderName(encoderIndex: SYSINT; var encoderName: WideString): WideString; dispid 204;
    function GetVideoQualityMaxValue(const deviceName: WideString; qualityType: TQualityType; 
                                     var maxValue: SYSINT): WideString; dispid 205;
    function GetVideoAnalogCount(var analogCount: SYSINT): WideString; dispid 206;
    function GetVideoAnalogName(analogIndex: SYSINT; var analogName: WideString): WideString; dispid 207;
    function GetVideoSizeCount(var sizeCount: SYSINT): WideString; dispid 208;
    function GetVideoSizeName(sizeIndex: SYSINT; var sizeName: WideString): WideString; dispid 209;
    function GetVideoColorDepthCount(var colorDepthCount: SYSINT): WideString; dispid 210;
    function GetVideoColorDepth(colorDepthIndex: SYSINT; var colorDepth: SYSINT): WideString; dispid 211;
    function CheckIsVfwDevice(const deviceName: WideString): WordBool; dispid 212;
    function CheckIsSupportVmr: WordBool; dispid 213;
    function VideoSizeConvert(const videoSize: WideString): {??TVideoSize}OleVariant; dispid 214;
    function GetIsSupportQuailtiCfg(const deviceName: WideString): WordBool; dispid 215;
  end;

// *********************************************************************//
// Interface: IDSPlay
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {36F986E0-6834-40BB-A444-18613D51FC10}
// *********************************************************************//
  IDSPlay = interface(IDispatch)
    ['{36F986E0-6834-40BB-A444-18613D51FC10}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_AutoScroll: WordBool; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    function Get_Color: OLE_COLOR; safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    function Get_Font: IFontDisp; safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    function Get_KeyPreview: WordBool; safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    function Get_Scaled: WordBool; safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    function Get_Active: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    function Get_HelpFile: WideString; safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    function Get_ScreenSnap: WordBool; safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    function Get_SnapBuffer: Integer; safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    function Play(const videoFile: WideString): WideString; safecall;
    function Pause: WideString; safecall;
    function Stop: WideString; safecall;
    function CaptureBmpImgToFile(const fileName: WideString): WideString; safecall;
    function CaptureJpgImgToFile(const fileName: WideString; compressRate: SYSINT): WideString; safecall;
    function AddRate: WideString; safecall;
    function DecRate: WideString; safecall;
    function RestoreRate: WideString; safecall;
    function ShowVideoInfo(parentHandle: SYSINT): WideString; safecall;
    procedure FreeRes; safecall;
    function Run: WideString; safecall;
    function FirstFrame: WideString; safecall;
    function LastFrame: WideString; safecall;
    function PriorFrame: WideString; safecall;
    function NextFrame: WideString; safecall;
    function Get_timeLen: SYSINT; safecall;
    function Get_FrameLen: SYSINT; safecall;
    function Get_CurTime: SYSINT; safecall;
    procedure Set_CurTime(Value: SYSINT); safecall;
    function Get_CurFrame: SYSINT; safecall;
    procedure Set_CurFrame(Value: SYSINT); safecall;
    function Get_PlayRate: Double; safecall;
    procedure Set_PlayRate(Value: Double); safecall;
    function Get_VideoState: TVideoState; safecall;
    function Get_ShowModel: TShowModel; safecall;
    procedure Set_ShowModel(Value: TShowModel); safecall;
    function Get_IsFullScreen: WordBool; safecall;
    procedure Set_IsFullScreen(Value: WordBool); safecall;
    function Get_IsFit: WordBool; safecall;
    procedure Set_IsFit(Value: WordBool); safecall;
    function Get_IsStretch: WordBool; safecall;
    procedure Set_IsStretch(Value: WordBool); safecall;
    function Get_IsAdjustWindowSize: WordBool; safecall;
    procedure Set_IsAdjustWindowSize(Value: WordBool); safecall;
    function Get_IsShowState: WordBool; safecall;
    procedure Set_IsShowState(Value: WordBool); safecall;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString; safecall;
    function QuitFullScreen: WideString; safecall;
    function RefreshWindow: WideString; safecall;
    function Get_IsEscKeyQuitFullScreen: WordBool; safecall;
    procedure Set_IsEscKeyQuitFullScreen(Value: WordBool); safecall;
    function Get_IsDblClickQuitFullScreen: WordBool; safecall;
    procedure Set_IsDblClickQuitFullScreen(Value: WordBool); safecall;
    function Get_IsClickQuitFullScreen: WordBool; safecall;
    procedure Set_IsClickQuitFullScreen(Value: WordBool); safecall;
    function GetVideoProperty(propertyType: TVideoProperty; var Value: WideString): WideString; safecall;
    function RePlay: WideString; safecall;
    function Get_CurWidth: Integer; safecall;
    procedure Set_CurWidth(Value: Integer); safecall;
    function Get_CurHeight: Integer; safecall;
    procedure Set_CurHeight(Value: Integer); safecall;
    function Get_SnatchWay: TSnatchWay; safecall;
    procedure Set_SnatchWay(Value: TSnatchWay); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function CaptureImgToClipBoard: WideString; safecall;
    function Get_Volume: Integer; safecall;
    procedure Set_Volume(Value: Integer); safecall;
    function Get_Balance: Integer; safecall;
    procedure Set_Balance(Value: Integer); safecall;
    function Get_StreamTypeName: WideString; safecall;
    procedure ShowAnimate(AnimateType: TAnimateType); safecall;
    procedure HideAnimate; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property AutoScroll: WordBool read Get_AutoScroll write Set_AutoScroll;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property AxBorderStyle: TxActiveFormBorderStyle read Get_AxBorderStyle write Set_AxBorderStyle;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Color: OLE_COLOR read Get_Color write Set_Color;
    property Font: IFontDisp read Get_Font write Set_Font;
    property KeyPreview: WordBool read Get_KeyPreview write Set_KeyPreview;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property PrintScale: TxPrintScale read Get_PrintScale write Set_PrintScale;
    property Scaled: WordBool read Get_Scaled write Set_Scaled;
    property Active: WordBool read Get_Active;
    property DropTarget: WordBool read Get_DropTarget write Set_DropTarget;
    property HelpFile: WideString read Get_HelpFile write Set_HelpFile;
    property ScreenSnap: WordBool read Get_ScreenSnap write Set_ScreenSnap;
    property SnapBuffer: Integer read Get_SnapBuffer write Set_SnapBuffer;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property AlignDisabled: WordBool read Get_AlignDisabled;
    property VisibleDockClientCount: Integer read Get_VisibleDockClientCount;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property timeLen: SYSINT read Get_timeLen;
    property FrameLen: SYSINT read Get_FrameLen;
    property CurTime: SYSINT read Get_CurTime write Set_CurTime;
    property CurFrame: SYSINT read Get_CurFrame write Set_CurFrame;
    property PlayRate: Double read Get_PlayRate write Set_PlayRate;
    property VideoState: TVideoState read Get_VideoState;
    property ShowModel: TShowModel read Get_ShowModel write Set_ShowModel;
    property IsFullScreen: WordBool read Get_IsFullScreen write Set_IsFullScreen;
    property IsFit: WordBool read Get_IsFit write Set_IsFit;
    property IsStretch: WordBool read Get_IsStretch write Set_IsStretch;
    property IsAdjustWindowSize: WordBool read Get_IsAdjustWindowSize write Set_IsAdjustWindowSize;
    property IsShowState: WordBool read Get_IsShowState write Set_IsShowState;
    property IsEscKeyQuitFullScreen: WordBool read Get_IsEscKeyQuitFullScreen write Set_IsEscKeyQuitFullScreen;
    property IsDblClickQuitFullScreen: WordBool read Get_IsDblClickQuitFullScreen write Set_IsDblClickQuitFullScreen;
    property IsClickQuitFullScreen: WordBool read Get_IsClickQuitFullScreen write Set_IsClickQuitFullScreen;
    property CurWidth: Integer read Get_CurWidth write Set_CurWidth;
    property CurHeight: Integer read Get_CurHeight write Set_CurHeight;
    property SnatchWay: TSnatchWay read Get_SnatchWay write Set_SnatchWay;
    property AppHandle: Integer read Get_AppHandle write Set_AppHandle;
    property Volume: Integer read Get_Volume write Set_Volume;
    property Balance: Integer read Get_Balance write Set_Balance;
    property StreamTypeName: WideString read Get_StreamTypeName;
  end;

// *********************************************************************//
// DispIntf:  IDSPlayDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {36F986E0-6834-40BB-A444-18613D51FC10}
// *********************************************************************//
  IDSPlayDisp = dispinterface
    ['{36F986E0-6834-40BB-A444-18613D51FC10}']
    property Visible: WordBool dispid 201;
    property AutoScroll: WordBool dispid 202;
    property AutoSize: WordBool dispid 203;
    property AxBorderStyle: TxActiveFormBorderStyle dispid 204;
    property Caption: WideString dispid -518;
    property Color: OLE_COLOR dispid -501;
    property Font: IFontDisp dispid -512;
    property KeyPreview: WordBool dispid 205;
    property PixelsPerInch: Integer dispid 206;
    property PrintScale: TxPrintScale dispid 207;
    property Scaled: WordBool dispid 208;
    property Active: WordBool readonly dispid 209;
    property DropTarget: WordBool dispid 210;
    property HelpFile: WideString dispid 211;
    property ScreenSnap: WordBool dispid 212;
    property SnapBuffer: Integer dispid 213;
    property DoubleBuffered: WordBool dispid 214;
    property AlignDisabled: WordBool readonly dispid 215;
    property VisibleDockClientCount: Integer readonly dispid 216;
    property Enabled: WordBool dispid -514;
    function Play(const videoFile: WideString): WideString; dispid 217;
    function Pause: WideString; dispid 218;
    function Stop: WideString; dispid 219;
    function CaptureBmpImgToFile(const fileName: WideString): WideString; dispid 220;
    function CaptureJpgImgToFile(const fileName: WideString; compressRate: SYSINT): WideString; dispid 221;
    function AddRate: WideString; dispid 222;
    function DecRate: WideString; dispid 223;
    function RestoreRate: WideString; dispid 224;
    function ShowVideoInfo(parentHandle: SYSINT): WideString; dispid 225;
    procedure FreeRes; dispid 226;
    function Run: WideString; dispid 227;
    function FirstFrame: WideString; dispid 228;
    function LastFrame: WideString; dispid 229;
    function PriorFrame: WideString; dispid 230;
    function NextFrame: WideString; dispid 231;
    property timeLen: SYSINT readonly dispid 232;
    property FrameLen: SYSINT readonly dispid 233;
    property CurTime: SYSINT dispid 234;
    property CurFrame: SYSINT dispid 235;
    property PlayRate: Double dispid 236;
    property VideoState: TVideoState readonly dispid 237;
    property ShowModel: TShowModel dispid 238;
    property IsFullScreen: WordBool dispid 239;
    property IsFit: WordBool dispid 240;
    property IsStretch: WordBool dispid 241;
    property IsAdjustWindowSize: WordBool dispid 242;
    property IsShowState: WordBool dispid 243;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString; dispid 244;
    function QuitFullScreen: WideString; dispid 245;
    function RefreshWindow: WideString; dispid 246;
    property IsEscKeyQuitFullScreen: WordBool dispid 247;
    property IsDblClickQuitFullScreen: WordBool dispid 248;
    property IsClickQuitFullScreen: WordBool dispid 249;
    function GetVideoProperty(propertyType: TVideoProperty; var Value: WideString): WideString; dispid 250;
    function RePlay: WideString; dispid 251;
    property CurWidth: Integer dispid 252;
    property CurHeight: Integer dispid 253;
    property SnatchWay: TSnatchWay dispid 254;
    property AppHandle: Integer dispid 255;
    function CaptureImgToClipBoard: WideString; dispid 256;
    property Volume: Integer dispid 257;
    property Balance: Integer dispid 258;
    property StreamTypeName: WideString readonly dispid 259;
    procedure ShowAnimate(AnimateType: TAnimateType); dispid 260;
    procedure HideAnimate; dispid 261;
  end;

// *********************************************************************//
// DispIntf:  IDSPlayEvents
// Flags:     (4096) Dispatchable
// GUID:      {932BAAE5-451C-47B3-BD8E-43DF7C4EF698}
// *********************************************************************//
  IDSPlayEvents = dispinterface
    ['{932BAAE5-451C-47B3-BD8E-43DF7C4EF698}']
    procedure OnActivate; dispid 201;
    procedure OnClick; dispid 202;
    procedure OnCreate; dispid 203;
    procedure OnDblClick; dispid 204;
    procedure OnDestroy; dispid 205;
    procedure OnDeactivate; dispid 206;
    procedure OnKeyPress(var Key: Smallint); dispid 207;
    procedure OnPaint; dispid 208;
    procedure OnMouseDown(button: SYSINT; shift: SYSINT; x: SYSINT; y: SYSINT); dispid 209;
    procedure OnMouseMove(shift: SYSINT; x: SYSINT; y: SYSINT); dispid 210;
    procedure OnMouseUp(button: SYSINT; shift: SYSINT; x: SYSINT; y: SYSINT); dispid 211;
    procedure OnKeyDown(var Key: SYSINT; shift: SYSINT); dispid 212;
    procedure OnKeyUp(var Key: SYSINT; shift: SYSINT); dispid 213;
    procedure OnResize; dispid 214;
    procedure OnGotFocus; dispid 215;
    procedure OnLostFocus; dispid 216;
    procedure OnMouseWheel(shift: SYSINT; wheelDelta: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                           var handled: WordBool); dispid 217;
    procedure OnMouseWheelDown(shift: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                               var handled: WordBool); dispid 218;
    procedure OnMouseWheelUp(shift: SYSINT; mousePosX: SYSINT; mousePosY: SYSINT; 
                             var handled: WordBool); dispid 219;
  end;

// *********************************************************************//
// Interface: ITMCIAudio
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {15923E95-8673-41A8-904E-9E12EFEC925C}
// *********************************************************************//
  ITMCIAudio = interface(IDispatch)
    ['{15923E95-8673-41A8-904E-9E12EFEC925C}']
    function StartRecord: WordBool; safecall;
    procedure StopRecord; safecall;
    procedure PauseRecord; safecall;
    procedure RestartRecord; safecall;
    function Get_LineColor: OLE_COLOR; safecall;
    procedure Set_LineColor(Value: OLE_COLOR); safecall;
    function Get_MaxColor: OLE_COLOR; safecall;
    procedure Set_MaxColor(Value: OLE_COLOR); safecall;
    function Get_SampleCount: Integer; safecall;
    procedure Set_SampleCount(Value: Integer); safecall;
    function Get_DrawFrequency: Integer; safecall;
    procedure Set_DrawFrequency(Value: Integer); safecall;
    function Get_BackColor: OLE_COLOR; safecall;
    procedure Set_BackColor(Value: OLE_COLOR); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure InitiateAction; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure SetSubComponent(IsSubComponent: WordBool); safecall;
    function Get_Channels: TAxChannels; safecall;
    procedure Set_Channels(Value: TAxChannels); safecall;
    function Get_BitsPerSample: TAxBPS; safecall;
    procedure Set_BitsPerSample(Value: TAxBPS); safecall;
    function Get_SampleRate: Integer; safecall;
    procedure Set_SampleRate(Value: Integer); safecall;
    function Get_NoSamples: Integer; safecall;
    procedure Set_NoSamples(Value: Integer); safecall;
    function Get_SplitChannels: WordBool; safecall;
    procedure Set_SplitChannels(Value: WordBool); safecall;
    function Get_TrigLevel: Integer; safecall;
    procedure Set_TrigLevel(Value: Integer); safecall;
    function Get_Triggered: WordBool; safecall;
    procedure Set_Triggered(Value: WordBool); safecall;
    function Get_RecordFile: WideString; safecall;
    procedure Set_RecordFile(const Value: WideString); safecall;
    function Get_RecordSize: Double; safecall;
    function Get_RecordCurTime: Integer; safecall;
    procedure Set_RecordCurTime(Value: Integer); safecall;
    function Get_RecordPostion: Double; safecall;
    procedure Set_RecordPostion(Value: Double); safecall;
    function Get_AudioDeviceId: Integer; safecall;
    procedure Set_AudioDeviceId(Value: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(Value: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(Value: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(Value: Integer); safecall;
    function Get_Handle: OLE_HANDLE; safecall;
    function Get_Hint: WideString; safecall;
    procedure Set_Hint(const Value: WideString); safecall;
    function Get_ShowHint: WordBool; safecall;
    procedure Set_ShowHint(Value: WordBool); safecall;
    function Get_RecordState: TAxMCIAudioState; safecall;
    function Get_RecordTimeLen: Integer; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure FreeRes; safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function Get_BufferCount: TAxBufferCount; safecall;
    procedure Set_BufferCount(Value: TAxBufferCount); safecall;
    function Get_RecordInputCount: Integer; safecall;
    function Get_RecordInputName(index: Integer): WideString; safecall;
    procedure ShowFormatDialog; safecall;
    function Get_FormatTag: Integer; safecall;
    procedure Set_FormatTag(Value: Integer); safecall;
    function Get_IsCompressWav: WordBool; safecall;
    procedure Set_IsCompressWav(Value: WordBool); safecall;
    function Get_CompRate: Integer; safecall;
    procedure Set_CompRate(Value: Integer); safecall;
    function Get_ErrorMsg: WideString; safecall;
    property LineColor: OLE_COLOR read Get_LineColor write Set_LineColor;
    property MaxColor: OLE_COLOR read Get_MaxColor write Set_MaxColor;
    property SampleCount: Integer read Get_SampleCount write Set_SampleCount;
    property DrawFrequency: Integer read Get_DrawFrequency write Set_DrawFrequency;
    property BackColor: OLE_COLOR read Get_BackColor write Set_BackColor;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Channels: TAxChannels read Get_Channels write Set_Channels;
    property BitsPerSample: TAxBPS read Get_BitsPerSample write Set_BitsPerSample;
    property SampleRate: Integer read Get_SampleRate write Set_SampleRate;
    property NoSamples: Integer read Get_NoSamples write Set_NoSamples;
    property SplitChannels: WordBool read Get_SplitChannels write Set_SplitChannels;
    property TrigLevel: Integer read Get_TrigLevel write Set_TrigLevel;
    property Triggered: WordBool read Get_Triggered write Set_Triggered;
    property RecordFile: WideString read Get_RecordFile write Set_RecordFile;
    property RecordSize: Double read Get_RecordSize;
    property RecordCurTime: Integer read Get_RecordCurTime write Set_RecordCurTime;
    property RecordPostion: Double read Get_RecordPostion write Set_RecordPostion;
    property AudioDeviceId: Integer read Get_AudioDeviceId write Set_AudioDeviceId;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Handle: OLE_HANDLE read Get_Handle;
    property Hint: WideString read Get_Hint write Set_Hint;
    property ShowHint: WordBool read Get_ShowHint write Set_ShowHint;
    property RecordState: TAxMCIAudioState read Get_RecordState;
    property RecordTimeLen: Integer read Get_RecordTimeLen;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property AppHandle: Integer read Get_AppHandle write Set_AppHandle;
    property Title: WideString read Get_Title write Set_Title;
    property BufferCount: TAxBufferCount read Get_BufferCount write Set_BufferCount;
    property RecordInputCount: Integer read Get_RecordInputCount;
    property RecordInputName[index: Integer]: WideString read Get_RecordInputName;
    property FormatTag: Integer read Get_FormatTag write Set_FormatTag;
    property IsCompressWav: WordBool read Get_IsCompressWav write Set_IsCompressWav;
    property CompRate: Integer read Get_CompRate write Set_CompRate;
    property ErrorMsg: WideString read Get_ErrorMsg;
  end;

// *********************************************************************//
// DispIntf:  ITMCIAudioDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {15923E95-8673-41A8-904E-9E12EFEC925C}
// *********************************************************************//
  ITMCIAudioDisp = dispinterface
    ['{15923E95-8673-41A8-904E-9E12EFEC925C}']
    function StartRecord: WordBool; dispid 201;
    procedure StopRecord; dispid 202;
    procedure PauseRecord; dispid 203;
    procedure RestartRecord; dispid 204;
    property LineColor: OLE_COLOR dispid 207;
    property MaxColor: OLE_COLOR dispid 208;
    property SampleCount: Integer dispid 209;
    property DrawFrequency: Integer dispid 210;
    property BackColor: OLE_COLOR dispid -501;
    property Enabled: WordBool dispid -514;
    procedure InitiateAction; dispid 215;
    property Visible: WordBool dispid 219;
    procedure SetSubComponent(IsSubComponent: WordBool); dispid 220;
    property Channels: TAxChannels dispid 221;
    property BitsPerSample: TAxBPS dispid 222;
    property SampleRate: Integer dispid 223;
    property NoSamples: Integer dispid 224;
    property SplitChannels: WordBool dispid 225;
    property TrigLevel: Integer dispid 226;
    property Triggered: WordBool dispid 227;
    property RecordFile: WideString dispid 228;
    property RecordSize: Double readonly dispid 230;
    property RecordCurTime: Integer dispid 229;
    property RecordPostion: Double dispid 231;
    property AudioDeviceId: Integer dispid 232;
    property Width: Integer dispid 233;
    property Height: Integer dispid 234;
    property Left: Integer dispid 235;
    property Top: Integer dispid 236;
    property Handle: OLE_HANDLE readonly dispid 237;
    property Hint: WideString dispid 238;
    property ShowHint: WordBool dispid 239;
    property RecordState: TAxMCIAudioState readonly dispid 206;
    property RecordTimeLen: Integer readonly dispid 205;
    property DoubleBuffered: WordBool dispid 211;
    procedure FreeRes; dispid 212;
    property AppHandle: Integer dispid 213;
    property Title: WideString dispid 214;
    property BufferCount: TAxBufferCount dispid 216;
    property RecordInputCount: Integer readonly dispid 217;
    property RecordInputName[index: Integer]: WideString readonly dispid 218;
    procedure ShowFormatDialog; dispid 240;
    property FormatTag: Integer dispid 241;
    property IsCompressWav: WordBool dispid 242;
    property CompRate: Integer dispid 243;
    property ErrorMsg: WideString readonly dispid 244;
  end;

// *********************************************************************//
// DispIntf:  ITMCIAudioEvents
// Flags:     (4096) Dispatchable
// GUID:      {B4593540-4604-4E09-90AA-9AB097805AD7}
// *********************************************************************//
  ITMCIAudioEvents = dispinterface
    ['{B4593540-4604-4E09-90AA-9AB097805AD7}']
  end;

// *********************************************************************//
// Interface: ITMCIPlayer
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B3C684B1-40D8-46FF-9429-BC07BEC99877}
// *********************************************************************//
  ITMCIPlayer = interface(IDispatch)
    ['{B3C684B1-40D8-46FF-9429-BC07BEC99877}']
    procedure StopPlay; safecall;
    procedure PausePlay; safecall;
    procedure RestartPlay; safecall;
    function PlayFile(const fileName: WideString; NoOfRepeats: Integer): WordBool; safecall;
    function Get_BackColor: OLE_COLOR; safecall;
    procedure Set_BackColor(Value: OLE_COLOR); safecall;
    function Get_BitsPerSample: TAxBPS; safecall;
    procedure Set_BitsPerSample(Value: TAxBPS); safecall;
    function Get_Channels: TAxChannels; safecall;
    procedure Set_Channels(Value: TAxChannels); safecall;
    function Get_SepCtrl: WordBool; safecall;
    procedure Set_SepCtrl(Value: WordBool); safecall;
    function Get_LineColor: OLE_COLOR; safecall;
    procedure Set_LineColor(Value: OLE_COLOR); safecall;
    function Get_MaxColor: OLE_COLOR; safecall;
    procedure Set_MaxColor(Value: OLE_COLOR); safecall;
    function Get_SampleCount: Integer; safecall;
    procedure Set_SampleCount(Value: Integer); safecall;
    function Get_DrawFrequency: Integer; safecall;
    procedure Set_DrawFrequency(Value: Integer); safecall;
    function Get_Title: WideString; safecall;
    procedure Set_Title(const Value: WideString); safecall;
    function Get_BufferCount: TAxBufferCount; safecall;
    procedure Set_BufferCount(Value: TAxBufferCount); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure InitiateAction; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_PlayDeviceId: Integer; safecall;
    procedure Set_PlayDeviceId(Value: Integer); safecall;
    function Get_SampleRate: Integer; safecall;
    procedure Set_SampleRate(Value: Integer); safecall;
    function Get_NoSamples: Integer; safecall;
    procedure Set_NoSamples(Value: Integer); safecall;
    function Get_PlaySize: Double; safecall;
    function Get_PlayPosition: Double; safecall;
    procedure Set_PlayPosition(Value: Double); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    procedure FreeRes; safecall;
    function Get_OutputDeviceName(index: Integer): WideString; safecall;
    function Get_Handle: OLE_HANDLE; safecall;
    function Get_Hint: WideString; safecall;
    procedure Set_Hint(const Value: WideString); safecall;
    function Get_ShowHint: WordBool; safecall;
    procedure Set_ShowHint(Value: WordBool); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(Value: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(Value: Integer); safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(Value: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(Value: Integer); safecall;
    procedure SetSubComponent(IsSubComponent: WordBool); safecall;
    procedure SetVolume(LeftVolume: Integer; RightVolume: Integer); safecall;
    procedure GetVolume(var LeftVolume: Integer; var RightVolume: Integer); safecall;
    function Get_PlayCurTime: Integer; safecall;
    procedure Set_PlayCurTime(Value: Integer); safecall;
    function Get_OutputDeviceCount: Integer; safecall;
    function Get_PlayTimeLen: Integer; safecall;
    function Get_PlayState: TAxMCIAudioState; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure ShowFormatDialog; safecall;
    function Get_FormatTag: Integer; safecall;
    procedure Set_FormatTag(Value: Integer); safecall;
    function Get_ErrorMsg: WideString; safecall;
    property BackColor: OLE_COLOR read Get_BackColor write Set_BackColor;
    property BitsPerSample: TAxBPS read Get_BitsPerSample write Set_BitsPerSample;
    property Channels: TAxChannels read Get_Channels write Set_Channels;
    property SepCtrl: WordBool read Get_SepCtrl write Set_SepCtrl;
    property LineColor: OLE_COLOR read Get_LineColor write Set_LineColor;
    property MaxColor: OLE_COLOR read Get_MaxColor write Set_MaxColor;
    property SampleCount: Integer read Get_SampleCount write Set_SampleCount;
    property DrawFrequency: Integer read Get_DrawFrequency write Set_DrawFrequency;
    property Title: WideString read Get_Title write Set_Title;
    property BufferCount: TAxBufferCount read Get_BufferCount write Set_BufferCount;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property PlayDeviceId: Integer read Get_PlayDeviceId write Set_PlayDeviceId;
    property SampleRate: Integer read Get_SampleRate write Set_SampleRate;
    property NoSamples: Integer read Get_NoSamples write Set_NoSamples;
    property PlaySize: Double read Get_PlaySize;
    property PlayPosition: Double read Get_PlayPosition write Set_PlayPosition;
    property AppHandle: Integer read Get_AppHandle write Set_AppHandle;
    property OutputDeviceName[index: Integer]: WideString read Get_OutputDeviceName;
    property Handle: OLE_HANDLE read Get_Handle;
    property Hint: WideString read Get_Hint write Set_Hint;
    property ShowHint: WordBool read Get_ShowHint write Set_ShowHint;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property PlayCurTime: Integer read Get_PlayCurTime write Set_PlayCurTime;
    property OutputDeviceCount: Integer read Get_OutputDeviceCount;
    property PlayTimeLen: Integer read Get_PlayTimeLen;
    property PlayState: TAxMCIAudioState read Get_PlayState;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property FormatTag: Integer read Get_FormatTag write Set_FormatTag;
    property ErrorMsg: WideString read Get_ErrorMsg;
  end;

// *********************************************************************//
// DispIntf:  ITMCIPlayerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B3C684B1-40D8-46FF-9429-BC07BEC99877}
// *********************************************************************//
  ITMCIPlayerDisp = dispinterface
    ['{B3C684B1-40D8-46FF-9429-BC07BEC99877}']
    procedure StopPlay; dispid 203;
    procedure PausePlay; dispid 204;
    procedure RestartPlay; dispid 205;
    function PlayFile(const fileName: WideString; NoOfRepeats: Integer): WordBool; dispid 206;
    property BackColor: OLE_COLOR dispid -501;
    property BitsPerSample: TAxBPS dispid 211;
    property Channels: TAxChannels dispid 212;
    property SepCtrl: WordBool dispid 215;
    property LineColor: OLE_COLOR dispid 217;
    property MaxColor: OLE_COLOR dispid 218;
    property SampleCount: Integer dispid 219;
    property DrawFrequency: Integer dispid 220;
    property Title: WideString dispid 221;
    property BufferCount: TAxBufferCount dispid 222;
    property Enabled: WordBool dispid -514;
    procedure InitiateAction; dispid 227;
    property Visible: WordBool dispid 231;
    property PlayDeviceId: Integer dispid 201;
    property SampleRate: Integer dispid 202;
    property NoSamples: Integer dispid 213;
    property PlaySize: Double readonly dispid 214;
    property PlayPosition: Double dispid 216;
    property AppHandle: Integer dispid 224;
    procedure FreeRes; dispid 225;
    property OutputDeviceName[index: Integer]: WideString readonly dispid 226;
    property Handle: OLE_HANDLE readonly dispid 228;
    property Hint: WideString dispid 229;
    property ShowHint: WordBool dispid 230;
    property Width: Integer dispid 232;
    property Height: Integer dispid 233;
    property Left: Integer dispid 234;
    property Top: Integer dispid 235;
    procedure SetSubComponent(IsSubComponent: WordBool); dispid 236;
    procedure SetVolume(LeftVolume: Integer; RightVolume: Integer); dispid 237;
    procedure GetVolume(var LeftVolume: Integer; var RightVolume: Integer); dispid 238;
    property PlayCurTime: Integer dispid 207;
    property OutputDeviceCount: Integer readonly dispid 209;
    property PlayTimeLen: Integer readonly dispid 208;
    property PlayState: TAxMCIAudioState readonly dispid 210;
    property DoubleBuffered: WordBool dispid 223;
    procedure ShowFormatDialog; dispid 239;
    property FormatTag: Integer dispid 240;
    property ErrorMsg: WideString readonly dispid 241;
  end;

// *********************************************************************//
// DispIntf:  ITMCIPlayerEvents
// Flags:     (4096) Dispatchable
// GUID:      {926B0355-E7BD-4277-85BF-FAA7DBF10133}
// *********************************************************************//
  ITMCIPlayerEvents = dispinterface
    ['{926B0355-E7BD-4277-85BF-FAA7DBF10133}']
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TDSPlay
// Help String      : DSPlay Control
// Default Interface: IDSPlay
// Def. Intf. DISP? : No
// Event   Interface: IDSPlayEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TDSPlayOnKeyPress = procedure(ASender: TObject; var Key: Smallint) of object;
  TDSPlayOnMouseDown = procedure(ASender: TObject; button: SYSINT; shift: SYSINT; x: SYSINT; 
                                                   y: SYSINT) of object;
  TDSPlayOnMouseMove = procedure(ASender: TObject; shift: SYSINT; x: SYSINT; y: SYSINT) of object;
  TDSPlayOnMouseUp = procedure(ASender: TObject; button: SYSINT; shift: SYSINT; x: SYSINT; y: SYSINT) of object;
  TDSPlayOnKeyDown = procedure(ASender: TObject; var Key: SYSINT; shift: SYSINT) of object;
  TDSPlayOnKeyUp = procedure(ASender: TObject; var Key: SYSINT; shift: SYSINT) of object;
  TDSPlayOnMouseWheel = procedure(ASender: TObject; shift: SYSINT; wheelDelta: SYSINT; 
                                                    mousePosX: SYSINT; mousePosY: SYSINT; 
                                                    var handled: WordBool) of object;
  TDSPlayOnMouseWheelDown = procedure(ASender: TObject; shift: SYSINT; mousePosX: SYSINT; 
                                                        mousePosY: SYSINT; var handled: WordBool) of object;
  TDSPlayOnMouseWheelUp = procedure(ASender: TObject; shift: SYSINT; mousePosX: SYSINT; 
                                                      mousePosY: SYSINT; var handled: WordBool) of object;

  TDSPlay = class(TOleControl)
  private
    FOnActivate: TNotifyEvent;
    FOnClick: TNotifyEvent;
    FOnCreate: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnDestroy: TNotifyEvent;
    FOnDeactivate: TNotifyEvent;
    FOnKeyPress: TDSPlayOnKeyPress;
    FOnPaint: TNotifyEvent;
    FOnMouseDown: TDSPlayOnMouseDown;
    FOnMouseMove: TDSPlayOnMouseMove;
    FOnMouseUp: TDSPlayOnMouseUp;
    FOnKeyDown: TDSPlayOnKeyDown;
    FOnKeyUp: TDSPlayOnKeyUp;
    FOnResize: TNotifyEvent;
    FOnGotFocus: TNotifyEvent;
    FOnLostFocus: TNotifyEvent;
    FOnMouseWheel: TDSPlayOnMouseWheel;
    FOnMouseWheelDown: TDSPlayOnMouseWheelDown;
    FOnMouseWheelUp: TDSPlayOnMouseWheelUp;
    FIntf: IDSPlay;
    function  GetControlInterface: IDSPlay;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function Play(const videoFile: WideString): WideString;
    function Pause: WideString;
    function Stop: WideString;
    function CaptureBmpImgToFile(const fileName: WideString): WideString;
    function CaptureJpgImgToFile(const fileName: WideString; compressRate: SYSINT): WideString;
    function AddRate: WideString;
    function DecRate: WideString;
    function RestoreRate: WideString;
    function ShowVideoInfo(parentHandle: SYSINT): WideString;
    procedure FreeRes;
    function Run: WideString;
    function FirstFrame: WideString;
    function LastFrame: WideString;
    function PriorFrame: WideString;
    function NextFrame: WideString;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString;
    function QuitFullScreen: WideString;
    function RefreshWindow: WideString;
    function GetVideoProperty(propertyType: TVideoProperty; var Value: WideString): WideString;
    function RePlay: WideString;
    function CaptureImgToClipBoard: WideString;
    procedure ShowAnimate(AnimateType: TAnimateType);
    procedure HideAnimate;
    property  ControlInterface: IDSPlay read GetControlInterface;
    property  DefaultInterface: IDSPlay read GetControlInterface;
    property Visible: WordBool index 201 read GetWordBoolProp write SetWordBoolProp;
    property Active: WordBool index 209 read GetWordBoolProp;
    property DropTarget: WordBool index 210 read GetWordBoolProp write SetWordBoolProp;
    property HelpFile: WideString index 211 read GetWideStringProp write SetWideStringProp;
    property ScreenSnap: WordBool index 212 read GetWordBoolProp write SetWordBoolProp;
    property SnapBuffer: Integer index 213 read GetIntegerProp write SetIntegerProp;
    property DoubleBuffered: WordBool index 214 read GetWordBoolProp write SetWordBoolProp;
    property AlignDisabled: WordBool index 215 read GetWordBoolProp;
    property VisibleDockClientCount: Integer index 216 read GetIntegerProp;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
    property timeLen: Integer index 232 read GetIntegerProp;
    property FrameLen: Integer index 233 read GetIntegerProp;
    property VideoState: TOleEnum index 237 read GetTOleEnumProp;
    property StreamTypeName: WideString index 259 read GetWideStringProp;
  published
    property Anchors;
    property  ParentColor;
    property  ParentFont;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property AutoScroll: WordBool index 202 read GetWordBoolProp write SetWordBoolProp stored False;
    property AutoSize: WordBool index 203 read GetWordBoolProp write SetWordBoolProp stored False;
    property AxBorderStyle: TOleEnum index 204 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Caption: WideString index -518 read GetWideStringProp write SetWideStringProp stored False;
    property Color: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property Font: TFont index -512 read GetTFontProp write SetTFontProp stored False;
    property KeyPreview: WordBool index 205 read GetWordBoolProp write SetWordBoolProp stored False;
    property PixelsPerInch: Integer index 206 read GetIntegerProp write SetIntegerProp stored False;
    property PrintScale: TOleEnum index 207 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Scaled: WordBool index 208 read GetWordBoolProp write SetWordBoolProp stored False;
    property CurTime: Integer index 234 read GetIntegerProp write SetIntegerProp stored False;
    property CurFrame: Integer index 235 read GetIntegerProp write SetIntegerProp stored False;
    property PlayRate: Double index 236 read GetDoubleProp write SetDoubleProp stored False;
    property ShowModel: TOleEnum index 238 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property IsFullScreen: WordBool index 239 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsFit: WordBool index 240 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsStretch: WordBool index 241 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsAdjustWindowSize: WordBool index 242 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsShowState: WordBool index 243 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsEscKeyQuitFullScreen: WordBool index 247 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsDblClickQuitFullScreen: WordBool index 248 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsClickQuitFullScreen: WordBool index 249 read GetWordBoolProp write SetWordBoolProp stored False;
    property CurWidth: Integer index 252 read GetIntegerProp write SetIntegerProp stored False;
    property CurHeight: Integer index 253 read GetIntegerProp write SetIntegerProp stored False;
    property SnatchWay: TOleEnum index 254 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property AppHandle: Integer index 255 read GetIntegerProp write SetIntegerProp stored False;
    property Volume: Integer index 257 read GetIntegerProp write SetIntegerProp stored False;
    property Balance: Integer index 258 read GetIntegerProp write SetIntegerProp stored False;
    property OnActivate: TNotifyEvent read FOnActivate write FOnActivate;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnCreate: TNotifyEvent read FOnCreate write FOnCreate;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
    property OnDeactivate: TNotifyEvent read FOnDeactivate write FOnDeactivate;
    property OnKeyPress: TDSPlayOnKeyPress read FOnKeyPress write FOnKeyPress;
    property OnPaint: TNotifyEvent read FOnPaint write FOnPaint;
    property OnMouseDown: TDSPlayOnMouseDown read FOnMouseDown write FOnMouseDown;
    property OnMouseMove: TDSPlayOnMouseMove read FOnMouseMove write FOnMouseMove;
    property OnMouseUp: TDSPlayOnMouseUp read FOnMouseUp write FOnMouseUp;
    property OnKeyDown: TDSPlayOnKeyDown read FOnKeyDown write FOnKeyDown;
    property OnKeyUp: TDSPlayOnKeyUp read FOnKeyUp write FOnKeyUp;
    property OnResize: TNotifyEvent read FOnResize write FOnResize;
    property OnGotFocus: TNotifyEvent read FOnGotFocus write FOnGotFocus;
    property OnLostFocus: TNotifyEvent read FOnLostFocus write FOnLostFocus;
    property OnMouseWheel: TDSPlayOnMouseWheel read FOnMouseWheel write FOnMouseWheel;
    property OnMouseWheelDown: TDSPlayOnMouseWheelDown read FOnMouseWheelDown write FOnMouseWheelDown;
    property OnMouseWheelUp: TDSPlayOnMouseWheelUp read FOnMouseWheelUp write FOnMouseWheelUp;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TDSCapture
// Help String      : DSCapture Control
// Default Interface: IDSCapture
// Def. Intf. DISP? : No
// Event   Interface: IDSCaptureEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TDSCaptureOnKeyPress = procedure(ASender: TObject; var Key: Smallint) of object;
  TDSCaptureOnMouseDown = procedure(ASender: TObject; button: SYSINT; shift: SYSINT; x: SYSINT; 
                                                      y: SYSINT) of object;
  TDSCaptureOnMouseMove = procedure(ASender: TObject; shift: SYSINT; x: SYSINT; y: SYSINT) of object;
  TDSCaptureOnMouseUp = procedure(ASender: TObject; button: SYSINT; shift: SYSINT; x: SYSINT; 
                                                    y: SYSINT) of object;
  TDSCaptureOnKeyDown = procedure(ASender: TObject; var Key: SYSINT; shift: SYSINT) of object;
  TDSCaptureOnKeyUp = procedure(ASender: TObject; var Key: SYSINT; shift: SYSINT) of object;
  TDSCaptureOnVideoSizeChange = procedure(ASender: TObject; videoWidth: SYSINT; 
                                                            videoHieght: SYSINT; 
                                                            windowWidth: SYSINT; 
                                                            windowHeight: SYSINT) of object;
  TDSCaptureOnMouseWheel = procedure(ASender: TObject; shift: SYSINT; wheelDelta: SYSINT; 
                                                       mousePosX: SYSINT; mousePosY: SYSINT; 
                                                       var handled: WordBool) of object;
  TDSCaptureOnMouseWheelDown = procedure(ASender: TObject; shift: SYSINT; mousePosX: SYSINT; 
                                                           mousePosY: SYSINT; var handled: WordBool) of object;
  TDSCaptureOnMouseWheelUp = procedure(ASender: TObject; shift: SYSINT; mousePosX: SYSINT; 
                                                         mousePosY: SYSINT; var handled: WordBool) of object;

  TDSCapture = class(TOleControl)
  private
    FOnActivate: TNotifyEvent;
    FOnClick: TNotifyEvent;
    FOnCreate: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnDestroy: TNotifyEvent;
    FOnDeactivate: TNotifyEvent;
    FOnKeyPress: TDSCaptureOnKeyPress;
    FOnPaint: TNotifyEvent;
    FOnMouseDown: TDSCaptureOnMouseDown;
    FOnMouseMove: TDSCaptureOnMouseMove;
    FOnMouseUp: TDSCaptureOnMouseUp;
    FOnKeyDown: TDSCaptureOnKeyDown;
    FOnKeyUp: TDSCaptureOnKeyUp;
    FOnResize: TNotifyEvent;
    FOnGotFocus: TNotifyEvent;
    FOnLostFocus: TNotifyEvent;
    FOnVideoSizeChange: TDSCaptureOnVideoSizeChange;
    FOnMouseWheel: TDSCaptureOnMouseWheel;
    FOnMouseWheelDown: TDSCaptureOnMouseWheelDown;
    FOnMouseWheelUp: TDSCaptureOnMouseWheelUp;
    FIntf: IDSCapture;
    function  GetControlInterface: IDSCapture;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString;
    function StartPreview: WideString;
    procedure FreeRes;
    function CaptureBmpImageToFile(const fileName: WideString): WideString;
    function StartCaptureVideo(const fileName: WideString): WideString;
    function StopCaptureVideo(out videoFile: WideString): WideString;
    function StopPreview: WideString;
    function ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString;
    function ShowVideoCapturePinCfg(parentHandle: Integer): WideString;
    function ShowVfwVideoSourceCfg(parentHandle: Integer): WideString;
    function ShowVfwVideoFormatCfg(parentHandle: Integer): WideString;
    function ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString;
    function ReadParameterFromFile: WideString;
    function CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString;
    function RefreshWindow: WideString;
    function ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString;
    function QuitFullScreen: WideString;
    function UpdateVideoQuailty: WideString;
    function SaveParameterToFile: WideString;
    function GetCaptureParameter(var parameter: TCaptureParameter): WideString;
    function SetCaptureParameter(var parameter: TCaptureParameter): WideString;
    function RePreview: WideString;
    function ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString;
    function CaptureImgToClipBoard: WideString;
    function ShowVfwCompressCfg(parentHandle: Integer): WideString;
    function ShowVideoCrossbarCfg(parentHandle: Integer): WideString;
    function CaptureBmpImage: IPictureDisp;
    function CaptureJpgImage(compressRate: Integer): IPictureDisp;
    function GetRealVideoSize: TVideoSize;
    property  ControlInterface: IDSCapture read GetControlInterface;
    property  DefaultInterface: IDSCapture read GetControlInterface;
    property Visible: WordBool index 201 read GetWordBoolProp write SetWordBoolProp;
    property Active: WordBool index 209 read GetWordBoolProp;
    property DropTarget: WordBool index 210 read GetWordBoolProp write SetWordBoolProp;
    property HelpFile: WideString index 211 read GetWideStringProp write SetWideStringProp;
    property ScreenSnap: WordBool index 212 read GetWordBoolProp write SetWordBoolProp;
    property SnapBuffer: Integer index 213 read GetIntegerProp write SetIntegerProp;
    property DoubleBuffered: WordBool index 214 read GetWordBoolProp write SetWordBoolProp;
    property AlignDisabled: WordBool index 215 read GetWordBoolProp;
    property VisibleDockClientCount: Integer index 216 read GetIntegerProp;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
    property PreviewState: WordBool index 234 read GetWordBoolProp;
    property CaptureState: WordBool index 235 read GetWordBoolProp;
  published
    property Anchors;
    property  ParentColor;
    property  ParentFont;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property AutoScroll: WordBool index 202 read GetWordBoolProp write SetWordBoolProp stored False;
    property AutoSize: WordBool index 203 read GetWordBoolProp write SetWordBoolProp stored False;
    property AxBorderStyle: TOleEnum index 204 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Caption: WideString index -518 read GetWideStringProp write SetWideStringProp stored False;
    property Color: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property Font: TFont index -512 read GetTFontProp write SetTFontProp stored False;
    property KeyPreview: WordBool index 205 read GetWordBoolProp write SetWordBoolProp stored False;
    property PixelsPerInch: Integer index 206 read GetIntegerProp write SetIntegerProp stored False;
    property PrintScale: TOleEnum index 207 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Scaled: WordBool index 208 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsStretch: WordBool index 223 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsShowState: WordBool index 224 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsFullScreen: WordBool index 225 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsAdjustWindowSize: WordBool index 226 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsFit: WordBool index 229 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsEscKeyQuitFullScreen: WordBool index 238 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsDblClickQuitFullScreen: WordBool index 239 read GetWordBoolProp write SetWordBoolProp stored False;
    property IsClickQuitFullScreen: WordBool index 240 read GetWordBoolProp write SetWordBoolProp stored False;
    property CurWidth: Integer index 241 read GetIntegerProp write SetIntegerProp stored False;
    property CurHeight: Integer index 242 read GetIntegerProp write SetIntegerProp stored False;
    property CurVideoWidth: Integer index 243 read GetIntegerProp write SetIntegerProp stored False;
    property CurVideoHeight: Integer index 244 read GetIntegerProp write SetIntegerProp stored False;
    property ShowModel: TOleEnum index 246 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property CapParameterWindPos: TOleEnum index 247 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property SnatchWay: TOleEnum index 250 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property ParameterCfgFileName: WideString index 257 read GetWideStringProp write SetWideStringProp stored False;
    property HideCfgItem: Integer index 258 read GetIntegerProp write SetIntegerProp stored False;
    property AppHandle: Integer index 259 read GetIntegerProp write SetIntegerProp stored False;
    property OnActivate: TNotifyEvent read FOnActivate write FOnActivate;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnCreate: TNotifyEvent read FOnCreate write FOnCreate;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
    property OnDeactivate: TNotifyEvent read FOnDeactivate write FOnDeactivate;
    property OnKeyPress: TDSCaptureOnKeyPress read FOnKeyPress write FOnKeyPress;
    property OnPaint: TNotifyEvent read FOnPaint write FOnPaint;
    property OnMouseDown: TDSCaptureOnMouseDown read FOnMouseDown write FOnMouseDown;
    property OnMouseMove: TDSCaptureOnMouseMove read FOnMouseMove write FOnMouseMove;
    property OnMouseUp: TDSCaptureOnMouseUp read FOnMouseUp write FOnMouseUp;
    property OnKeyDown: TDSCaptureOnKeyDown read FOnKeyDown write FOnKeyDown;
    property OnKeyUp: TDSCaptureOnKeyUp read FOnKeyUp write FOnKeyUp;
    property OnResize: TNotifyEvent read FOnResize write FOnResize;
    property OnGotFocus: TNotifyEvent read FOnGotFocus write FOnGotFocus;
    property OnLostFocus: TNotifyEvent read FOnLostFocus write FOnLostFocus;
    property OnVideoSizeChange: TDSCaptureOnVideoSizeChange read FOnVideoSizeChange write FOnVideoSizeChange;
    property OnMouseWheel: TDSCaptureOnMouseWheel read FOnMouseWheel write FOnMouseWheel;
    property OnMouseWheelDown: TDSCaptureOnMouseWheelDown read FOnMouseWheelDown write FOnMouseWheelDown;
    property OnMouseWheelUp: TDSCaptureOnMouseWheelUp read FOnMouseWheelUp write FOnMouseWheelUp;
  end;

// *********************************************************************//
// The Class CoDSCapParameterEnum provides a Create and CreateRemote method to          
// create instances of the default interface IDSParameterEnum exposed by              
// the CoClass DSCapParameterEnum. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDSCapParameterEnum = class
    class function Create: IDSParameterEnum;
    class function CreateRemote(const MachineName: string): IDSParameterEnum;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TTMCIAudio
// Help String      : TMCIAudio Control
// Default Interface: ITMCIAudio
// Def. Intf. DISP? : No
// Event   Interface: ITMCIAudioEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TTMCIAudio = class(TOleControl)
  private
    FIntf: ITMCIAudio;
    function  GetControlInterface: ITMCIAudio;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function Get_RecordInputName(index: Integer): WideString;
  public
    function StartRecord: WordBool;
    procedure StopRecord;
    procedure PauseRecord;
    procedure RestartRecord;
    procedure InitiateAction;
    procedure SetSubComponent(IsSubComponent: WordBool);
    procedure FreeRes;
    procedure ShowFormatDialog;
    property  ControlInterface: ITMCIAudio read GetControlInterface;
    property  DefaultInterface: ITMCIAudio read GetControlInterface;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
    property Visible: WordBool index 219 read GetWordBoolProp write SetWordBoolProp;
    property RecordSize: Double index 230 read GetDoubleProp;
    property Handle: Integer index 237 read GetIntegerProp;
    property RecordState: TOleEnum index 206 read GetTOleEnumProp;
    property RecordTimeLen: Integer index 205 read GetIntegerProp;
    property RecordInputCount: Integer index 217 read GetIntegerProp;
    property RecordInputName[index: Integer]: WideString read Get_RecordInputName;
    property ErrorMsg: WideString index 244 read GetWideStringProp;
  published
    property Anchors;
    property  ParentColor;
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property LineColor: TColor index 207 read GetTColorProp write SetTColorProp stored False;
    property MaxColor: TColor index 208 read GetTColorProp write SetTColorProp stored False;
    property SampleCount: Integer index 209 read GetIntegerProp write SetIntegerProp stored False;
    property DrawFrequency: Integer index 210 read GetIntegerProp write SetIntegerProp stored False;
    property BackColor: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property Channels: TOleEnum index 221 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property BitsPerSample: TOleEnum index 222 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property SampleRate: Integer index 223 read GetIntegerProp write SetIntegerProp stored False;
    property NoSamples: Integer index 224 read GetIntegerProp write SetIntegerProp stored False;
    property SplitChannels: WordBool index 225 read GetWordBoolProp write SetWordBoolProp stored False;
    property TrigLevel: Integer index 226 read GetIntegerProp write SetIntegerProp stored False;
    property Triggered: WordBool index 227 read GetWordBoolProp write SetWordBoolProp stored False;
    property RecordFile: WideString index 228 read GetWideStringProp write SetWideStringProp stored False;
    property RecordCurTime: Integer index 229 read GetIntegerProp write SetIntegerProp stored False;
    property RecordPostion: Double index 231 read GetDoubleProp write SetDoubleProp stored False;
    property AudioDeviceId: Integer index 232 read GetIntegerProp write SetIntegerProp stored False;
    property Hint: WideString index 238 read GetWideStringProp write SetWideStringProp stored False;
    property ShowHint: WordBool index 239 read GetWordBoolProp write SetWordBoolProp stored False;
    property DoubleBuffered: WordBool index 211 read GetWordBoolProp write SetWordBoolProp stored False;
    property AppHandle: Integer index 213 read GetIntegerProp write SetIntegerProp stored False;
    property Title: WideString index 214 read GetWideStringProp write SetWideStringProp stored False;
    property BufferCount: TOleEnum index 216 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property FormatTag: Integer index 241 read GetIntegerProp write SetIntegerProp stored False;
    property IsCompressWav: WordBool index 242 read GetWordBoolProp write SetWordBoolProp stored False;
    property CompRate: Integer index 243 read GetIntegerProp write SetIntegerProp stored False;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TTMCIPlayer
// Help String      : TMCIPlayer Control
// Default Interface: ITMCIPlayer
// Def. Intf. DISP? : No
// Event   Interface: ITMCIPlayerEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TTMCIPlayer = class(TOleControl)
  private
    FIntf: ITMCIPlayer;
    function  GetControlInterface: ITMCIPlayer;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function Get_OutputDeviceName(index: Integer): WideString;
  public
    procedure StopPlay;
    procedure PausePlay;
    procedure RestartPlay;
    function PlayFile(const fileName: WideString; NoOfRepeats: Integer): WordBool;
    procedure InitiateAction;
    procedure FreeRes;
    procedure SetSubComponent(IsSubComponent: WordBool);
    procedure SetVolume(LeftVolume: Integer; RightVolume: Integer);
    procedure GetVolume(var LeftVolume: Integer; var RightVolume: Integer);
    procedure ShowFormatDialog;
    property  ControlInterface: ITMCIPlayer read GetControlInterface;
    property  DefaultInterface: ITMCIPlayer read GetControlInterface;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
    property Visible: WordBool index 231 read GetWordBoolProp write SetWordBoolProp;
    property PlaySize: Double index 214 read GetDoubleProp;
    property OutputDeviceName[index: Integer]: WideString read Get_OutputDeviceName;
    property Handle: Integer index 228 read GetIntegerProp;
    property OutputDeviceCount: Integer index 209 read GetIntegerProp;
    property PlayTimeLen: Integer index 208 read GetIntegerProp;
    property PlayState: TOleEnum index 210 read GetTOleEnumProp;
    property ErrorMsg: WideString index 241 read GetWideStringProp;
  published
    property Anchors;
    property  ParentColor;
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property BackColor: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property BitsPerSample: TOleEnum index 211 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Channels: TOleEnum index 212 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property SepCtrl: WordBool index 215 read GetWordBoolProp write SetWordBoolProp stored False;
    property LineColor: TColor index 217 read GetTColorProp write SetTColorProp stored False;
    property MaxColor: TColor index 218 read GetTColorProp write SetTColorProp stored False;
    property SampleCount: Integer index 219 read GetIntegerProp write SetIntegerProp stored False;
    property DrawFrequency: Integer index 220 read GetIntegerProp write SetIntegerProp stored False;
    property Title: WideString index 221 read GetWideStringProp write SetWideStringProp stored False;
    property BufferCount: TOleEnum index 222 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property PlayDeviceId: Integer index 201 read GetIntegerProp write SetIntegerProp stored False;
    property SampleRate: Integer index 202 read GetIntegerProp write SetIntegerProp stored False;
    property NoSamples: Integer index 213 read GetIntegerProp write SetIntegerProp stored False;
    property PlayPosition: Double index 216 read GetDoubleProp write SetDoubleProp stored False;
    property AppHandle: Integer index 224 read GetIntegerProp write SetIntegerProp stored False;
    property Hint: WideString index 229 read GetWideStringProp write SetWideStringProp stored False;
    property ShowHint: WordBool index 230 read GetWordBoolProp write SetWordBoolProp stored False;
    property PlayCurTime: Integer index 207 read GetIntegerProp write SetIntegerProp stored False;
    property DoubleBuffered: WordBool index 223 read GetWordBoolProp write SetWordBoolProp stored False;
    property FormatTag: Integer index 240 read GetIntegerProp write SetIntegerProp stored False;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TDSPlay.InitControlData;
const
  CEventDispIDs: array [0..18] of DWORD = (
    $000000C9, $000000CA, $000000CB, $000000CC, $000000CD, $000000CE,
    $000000CF, $000000D0, $000000D1, $000000D2, $000000D3, $000000D4,
    $000000D5, $000000D6, $000000D7, $000000D8, $000000D9, $000000DA,
    $000000DB);
  CTFontIDs: array [0..0] of DWORD = (
    $FFFFFE00);
  CControlData: TControlData2 = (
    ClassID: '{BC410BFE-ED4B-4DFD-8506-2D6CB2BBF564}';
    EventIID: '{932BAAE5-451C-47B3-BD8E-43DF7C4EF698}';
    EventCount: 19;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $0000001D;
    Version: 401;
    FontCount: 1;
    FontIDs: @CTFontIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnActivate) - Cardinal(Self);
end;

procedure TDSPlay.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IDSPlay;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TDSPlay.GetControlInterface: IDSPlay;
begin
  CreateControl;
  Result := FIntf;
end;

function TDSPlay.Play(const videoFile: WideString): WideString;
begin
  Result := DefaultInterface.Play(videoFile);
end;

function TDSPlay.Pause: WideString;
begin
  Result := DefaultInterface.Pause;
end;

function TDSPlay.Stop: WideString;
begin
  Result := DefaultInterface.Stop;
end;

function TDSPlay.CaptureBmpImgToFile(const fileName: WideString): WideString;
begin
  Result := DefaultInterface.CaptureBmpImgToFile(fileName);
end;

function TDSPlay.CaptureJpgImgToFile(const fileName: WideString; compressRate: SYSINT): WideString;
begin
  Result := DefaultInterface.CaptureJpgImgToFile(fileName, compressRate);
end;

function TDSPlay.AddRate: WideString;
begin
  Result := DefaultInterface.AddRate;
end;

function TDSPlay.DecRate: WideString;
begin
  Result := DefaultInterface.DecRate;
end;

function TDSPlay.RestoreRate: WideString;
begin
  Result := DefaultInterface.RestoreRate;
end;

function TDSPlay.ShowVideoInfo(parentHandle: SYSINT): WideString;
begin
  Result := DefaultInterface.ShowVideoInfo(parentHandle);
end;

procedure TDSPlay.FreeRes;
begin
  DefaultInterface.FreeRes;
end;

function TDSPlay.Run: WideString;
begin
  Result := DefaultInterface.Run;
end;

function TDSPlay.FirstFrame: WideString;
begin
  Result := DefaultInterface.FirstFrame;
end;

function TDSPlay.LastFrame: WideString;
begin
  Result := DefaultInterface.LastFrame;
end;

function TDSPlay.PriorFrame: WideString;
begin
  Result := DefaultInterface.PriorFrame;
end;

function TDSPlay.NextFrame: WideString;
begin
  Result := DefaultInterface.NextFrame;
end;

function TDSPlay.ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString;
begin
  Result := DefaultInterface.ShowFullScreen(parentHandle, monitorIndex);
end;

function TDSPlay.QuitFullScreen: WideString;
begin
  Result := DefaultInterface.QuitFullScreen;
end;

function TDSPlay.RefreshWindow: WideString;
begin
  Result := DefaultInterface.RefreshWindow;
end;

function TDSPlay.GetVideoProperty(propertyType: TVideoProperty; var Value: WideString): WideString;
begin
  Result := DefaultInterface.GetVideoProperty(propertyType, Value);
end;

function TDSPlay.RePlay: WideString;
begin
  Result := DefaultInterface.RePlay;
end;

function TDSPlay.CaptureImgToClipBoard: WideString;
begin
  Result := DefaultInterface.CaptureImgToClipBoard;
end;

procedure TDSPlay.ShowAnimate(AnimateType: TAnimateType);
begin
  DefaultInterface.ShowAnimate(AnimateType);
end;

procedure TDSPlay.HideAnimate;
begin
  DefaultInterface.HideAnimate;
end;

procedure TDSCapture.InitControlData;
const
  CEventDispIDs: array [0..19] of DWORD = (
    $000000C9, $000000CA, $000000CB, $000000CC, $000000CD, $000000CE,
    $000000CF, $000000D0, $000000D1, $000000D2, $000000D3, $000000D4,
    $000000D5, $000000D6, $000000D7, $000000D8, $000000D9, $000000DA,
    $000000DB, $000000DC);
  CTFontIDs: array [0..0] of DWORD = (
    $FFFFFE00);
  CControlData: TControlData2 = (
    ClassID: '{137D6CFF-36DB-4AB2-BD2C-AC279626A8F3}';
    EventIID: '{EC14A323-3D09-443B-A23E-FD86909CD935}';
    EventCount: 20;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $0000001D;
    Version: 401;
    FontCount: 1;
    FontIDs: @CTFontIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnActivate) - Cardinal(Self);
end;

procedure TDSCapture.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IDSCapture;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TDSCapture.GetControlInterface: IDSCapture;
begin
  CreateControl;
  Result := FIntf;
end;

function TDSCapture.ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowCaptureParameterCfgDialog(parentHandle);
end;

function TDSCapture.StartPreview: WideString;
begin
  Result := DefaultInterface.StartPreview;
end;

procedure TDSCapture.FreeRes;
begin
  DefaultInterface.FreeRes;
end;

function TDSCapture.CaptureBmpImageToFile(const fileName: WideString): WideString;
begin
  Result := DefaultInterface.CaptureBmpImageToFile(fileName);
end;

function TDSCapture.StartCaptureVideo(const fileName: WideString): WideString;
begin
  Result := DefaultInterface.StartCaptureVideo(fileName);
end;

function TDSCapture.StopCaptureVideo(out videoFile: WideString): WideString;
begin
  Result := DefaultInterface.StopCaptureVideo(videoFile);
end;

function TDSCapture.StopPreview: WideString;
begin
  Result := DefaultInterface.StopPreview;
end;

function TDSCapture.ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVideoCaptureFilterCfg(parentHandle);
end;

function TDSCapture.ShowVideoCapturePinCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVideoCapturePinCfg(parentHandle);
end;

function TDSCapture.ShowVfwVideoSourceCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVfwVideoSourceCfg(parentHandle);
end;

function TDSCapture.ShowVfwVideoFormatCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVfwVideoFormatCfg(parentHandle);
end;

function TDSCapture.ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVfwVideoDisplayCfg(parentHandle);
end;

function TDSCapture.ReadParameterFromFile: WideString;
begin
  Result := DefaultInterface.ReadParameterFromFile;
end;

function TDSCapture.CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString;
begin
  Result := DefaultInterface.CaptureJpgImageToFile(fileName, compressRate);
end;

function TDSCapture.RefreshWindow: WideString;
begin
  Result := DefaultInterface.RefreshWindow;
end;

function TDSCapture.ShowFullScreen(parentHandle: Integer; monitorIndex: Integer): WideString;
begin
  Result := DefaultInterface.ShowFullScreen(parentHandle, monitorIndex);
end;

function TDSCapture.QuitFullScreen: WideString;
begin
  Result := DefaultInterface.QuitFullScreen;
end;

function TDSCapture.UpdateVideoQuailty: WideString;
begin
  Result := DefaultInterface.UpdateVideoQuailty;
end;

function TDSCapture.SaveParameterToFile: WideString;
begin
  Result := DefaultInterface.SaveParameterToFile;
end;

function TDSCapture.GetCaptureParameter(var parameter: TCaptureParameter): WideString;
begin
  Result := DefaultInterface.GetCaptureParameter(parameter);
end;

function TDSCapture.SetCaptureParameter(var parameter: TCaptureParameter): WideString;
begin
  Result := DefaultInterface.SetCaptureParameter(parameter);
end;

function TDSCapture.RePreview: WideString;
begin
  Result := DefaultInterface.RePreview;
end;

function TDSCapture.ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVideoEncoderFilterCfg(parentHandle);
end;

function TDSCapture.CaptureImgToClipBoard: WideString;
begin
  Result := DefaultInterface.CaptureImgToClipBoard;
end;

function TDSCapture.ShowVfwCompressCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVfwCompressCfg(parentHandle);
end;

function TDSCapture.ShowVideoCrossbarCfg(parentHandle: Integer): WideString;
begin
  Result := DefaultInterface.ShowVideoCrossbarCfg(parentHandle);
end;

function TDSCapture.CaptureBmpImage: IPictureDisp;
begin
  Result := DefaultInterface.CaptureBmpImage;
end;

function TDSCapture.CaptureJpgImage(compressRate: Integer): IPictureDisp;
begin
  Result := DefaultInterface.CaptureJpgImage(compressRate);
end;

function TDSCapture.GetRealVideoSize: TVideoSize;
begin
  Result := DefaultInterface.GetRealVideoSize;
end;

class function CoDSCapParameterEnum.Create: IDSParameterEnum;
begin
  Result := CreateComObject(CLASS_DSCapParameterEnum) as IDSParameterEnum;
end;

class function CoDSCapParameterEnum.CreateRemote(const MachineName: string): IDSParameterEnum;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DSCapParameterEnum) as IDSParameterEnum;
end;

procedure TTMCIAudio.InitControlData;
const
  CControlData: TControlData2 = (
    ClassID: '{E9F8B5F5-84D2-47CA-B2C0-ABC8E9B840A4}';
    EventIID: '';
    EventCount: 0;
    EventDispIDs: nil;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $00000009;
    Version: 401);
begin
  ControlData := @CControlData;
end;

procedure TTMCIAudio.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as ITMCIAudio;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TTMCIAudio.GetControlInterface: ITMCIAudio;
begin
  CreateControl;
  Result := FIntf;
end;

function TTMCIAudio.Get_RecordInputName(index: Integer): WideString;
begin
    Result := DefaultInterface.RecordInputName[index];
end;

function TTMCIAudio.StartRecord: WordBool;
begin
  Result := DefaultInterface.StartRecord;
end;

procedure TTMCIAudio.StopRecord;
begin
  DefaultInterface.StopRecord;
end;

procedure TTMCIAudio.PauseRecord;
begin
  DefaultInterface.PauseRecord;
end;

procedure TTMCIAudio.RestartRecord;
begin
  DefaultInterface.RestartRecord;
end;

procedure TTMCIAudio.InitiateAction;
begin
  DefaultInterface.InitiateAction;
end;

procedure TTMCIAudio.SetSubComponent(IsSubComponent: WordBool);
begin
  DefaultInterface.SetSubComponent(IsSubComponent);
end;

procedure TTMCIAudio.FreeRes;
begin
  DefaultInterface.FreeRes;
end;

procedure TTMCIAudio.ShowFormatDialog;
begin
  DefaultInterface.ShowFormatDialog;
end;

procedure TTMCIPlayer.InitControlData;
const
  CControlData: TControlData2 = (
    ClassID: '{F38977AD-8F88-4AE4-BD08-57584273CFF6}';
    EventIID: '';
    EventCount: 0;
    EventDispIDs: nil;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $00000009;
    Version: 401);
begin
  ControlData := @CControlData;
end;

procedure TTMCIPlayer.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as ITMCIPlayer;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TTMCIPlayer.GetControlInterface: ITMCIPlayer;
begin
  CreateControl;
  Result := FIntf;
end;

function TTMCIPlayer.Get_OutputDeviceName(index: Integer): WideString;
begin
    Result := DefaultInterface.OutputDeviceName[index];
end;

procedure TTMCIPlayer.StopPlay;
begin
  DefaultInterface.StopPlay;
end;

procedure TTMCIPlayer.PausePlay;
begin
  DefaultInterface.PausePlay;
end;

procedure TTMCIPlayer.RestartPlay;
begin
  DefaultInterface.RestartPlay;
end;

function TTMCIPlayer.PlayFile(const fileName: WideString; NoOfRepeats: Integer): WordBool;
begin
  Result := DefaultInterface.PlayFile(fileName, NoOfRepeats);
end;

procedure TTMCIPlayer.InitiateAction;
begin
  DefaultInterface.InitiateAction;
end;

procedure TTMCIPlayer.FreeRes;
begin
  DefaultInterface.FreeRes;
end;

procedure TTMCIPlayer.SetSubComponent(IsSubComponent: WordBool);
begin
  DefaultInterface.SetSubComponent(IsSubComponent);
end;

procedure TTMCIPlayer.SetVolume(LeftVolume: Integer; RightVolume: Integer);
begin
  DefaultInterface.SetVolume(LeftVolume, RightVolume);
end;

procedure TTMCIPlayer.GetVolume(var LeftVolume: Integer; var RightVolume: Integer);
begin
  DefaultInterface.GetVolume(LeftVolume, RightVolume);
end;

procedure TTMCIPlayer.ShowFormatDialog;
begin
  DefaultInterface.ShowFormatDialog;
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TDSPlay, TDSCapture, TTMCIAudio, TTMCIPlayer]);
end;

end.
