{*******************************************************************************
视频采集所用到的相关数据类型定义
创建人：TJH
创建日前：2009-11-3

描述：...


*******************************************************************************}
unit VideoProcessDefine;

interface

uses
  Classes, Windows, DirectShow9, ZLDSVideoProcess_TLB;

const
  CONST_TEMP_DIR: WideString = 'Temp\';
  //名义上，一个参考时钟以千万分之一秒的精度来度量时间，但是实际上的精度不会这么高
  //1秒=1000000微秒,1,000 微秒 = 1毫秒
  ONE_SECOND: Integer = 10000 * 1000;

  //一天的长度 86400000 毫秒
  MiliSecInOneDay = 86400000;

Type
  //旋转类型
  TCircumvolveType = (ctNormal{正常},
                      ctUprightness{垂直},
                      ctPlane{水平},
                      ctAngle{角度});

  //采集参数类型
  TCaptureParameterType = (cptCaptureDeviceName,
                           cptInputPinName,
                           cptOutputPinName,
                           cptVideoAnalog,
                           cptColorDepth,
                           cptVideoSize,
                           cptBrightness,
                           cptContrast,
                           cptHue,
                           cptSaturation,
                           cptGamma,
                           cptWhiteBlance,
                           cptEncoderName,
                           cptIsTimeLimit,
                           cptLimitLength,
                           cptIsConvert8Bit,
                           cptIsHintSound,
                           cptIsApplyImageCut,
                           cptTopRate,
                           cptHeightRate,
                           cptLeftRate,
                           cptWidthRate,
                           cptVideoShowModel,
                           cptSnatchWay,
                           cptIsShowState,
                           cptInputCrossbar,
                           cptOutputCrossbar,
                           cptIsAutoBrightness,
                           cptIsAutoContrast,
                           cptIsAutoHue,
                           cptIsAutoSaturation,
                           cptIsAutoGamma,
                           cptIsAutoWhiteBlance,
                           cptIsSoundHint
                           );

  //vfw配置类型                           
  TVfwConfigType = (vctVideoSourceProperty{视频源filter属性},
                    vctVideoCapturePinProperty{视频采集端口属性},
                    vctVfwVideoFormat{vfw视频格式设置},
                    vctVfwVideoDisplay{vfw视频显示设置},
                    vctVideoCrossbar{video Crossbar设置},
                    vctVfwCompressDialog{压缩对话框});


  //视频格式
  {TVideoSize = record
    Width: Integer;         //宽度
    Height: Integer;        //高度
  end;}

  TOtherPar = packed record
    Frame: Integer;
  end;
  
  //采集参数配置对象*******************************************
  TCaptureParameterConfig = class(TObject)
  private
  public
    //转换成标准制式定义
    class function ConvertAnalogVideoStandard(const analogName: WideString): Integer;
    //videoFormatStr格式:“320 X 240”
    class function ConvertVideoSizeInf(const videoSizeStr: WideString): TVideoSize;
    //复制参数到指定对象
    class procedure CopyParameter(const sourceParameter: TCaptureParameter;
      var destParameter: TCaptureParameter);
    //初始化采集参数
    class procedure InitCaptureParameter(var captureParameter: TCaptureParameter);
    //取得分辨率大小字符
    class function GetVideoSizeStr(const width, height: Integer): WideString;
  end;

  //参数改变事件
  TOnParameterChangeEvent = procedure(const parameter: TCaptureParameter; const needCaptureSample: Boolean) of object;
  //vfw配置事件
  TOnVfwConfigCallEvent = procedure(const operVfwConfigType: TVfwConfigType; const parentHandle: Integer; out errMsg: WideString) of object;


  //视频信息对象
  TVideoInf = class(TObject)
  public
    videoFile: WideString;        //视频文件路径全名
    MajorTypeName: WideString;    //主媒体类型名称
    SubTypeName: WideString;      //子媒体类型名称
    FormatTypeName: WideString;   //格式类型名称
    TimeFormatName: WideString;   //时间格式名称
    VideoColorDepth: Integer;     //视频颜色深度
    VideoWidth: Integer;          //视频宽度
    VideoHeight: Integer;         //视频高度
    StreamCount: Integer;         //流数量
    FrameRate: Double;            //帧速率
    TimeLen: Int64;               //时间长度(单位：秒)
    FrameLen: Int64;              //帧长度(单位：帧)
  end;

var
  SysVideoSize: array[0..11] of String=('160X120', '176X144', '240X180', '320X240', '352X288',
                                        '512X380', '640X480', '704X576', '720X576', '768X576', '800X600', '1024X768');
                                        
  SysVideoAnalog: array[0..22] of String=('PAL_B', 'PAL_D', 'PAL_G', 'PAL_H', 'PAL_I', 'PAL_M', 'PAL_N', 'PAL_60', 'PAL_Mask',
                                          'NTSC_M', 'NTSC_M_J', 'NTSC_433', 'NTSC_Mask',
                                          'SECAM_B', 'SECAM_D', 'SECAM_G', 'SECAM_H', 'SECAM_K', 'SECAM_K1', 'SECAM_L', 'SECAM_L1', 'SECAM_Mask',
                                          'None');

implementation

uses SysUtils, StrUtils, CaptureDebug;


const
  VIDEO_FORMAT_SPLIT: String = 'X';

{ TCaptureParameterConfig }

class function TCaptureParameterConfig.ConvertAnalogVideoStandard(const analogName: WideString): Integer;
begin
  if UpperCase(analogName) = UpperCase('None') then begin
    Result := AnalogVideo_None;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_B') then begin
    Result := AnalogVideo_PAL_B;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_D') then begin
    Result := AnalogVideo_PAL_D;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_G') then begin
    Result := AnalogVideo_PAL_G;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_H') then begin
    Result := AnalogVideo_PAL_H;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_I') then begin
    Result := AnalogVideo_PAL_I;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_M') then begin
    Result := AnalogVideo_PAL_M;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_N') then begin
    Result := AnalogVideo_PAL_N;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_60') then begin
    Result := AnalogVideo_PAL_60;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('PAL_Mask') then begin
    Result := AnalogVideo_PAL_Mask;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('NTSC_M') then begin
    Result := AnalogVideo_NTSC_M;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('NTSC_M_J') then begin
    Result := AnalogVideo_NTSC_M_J;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('NTSC_433') then begin
    Result := AnalogVideo_NTSC_433;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('NTSC_Mask') then begin
    Result := AnalogVideo_NTSC_Mask;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_B') then begin
    Result := AnalogVideo_SECAM_B;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_D') then begin
    Result := AnalogVideo_SECAM_D;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_G') then begin
    Result := AnalogVideo_SECAM_G;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_H') then begin
    Result := AnalogVideo_SECAM_H;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_K') then begin
    Result := AnalogVideo_SECAM_K;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_K1') then begin
    Result := AnalogVideo_SECAM_K1;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_L') then begin
    Result := AnalogVideo_SECAM_L;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_L1') then begin
    Result := AnalogVideo_SECAM_L1;
    Exit;
  end;

  if UpperCase(analogName) = UpperCase('SECAM_Mask') then begin
    Result := AnalogVideo_SECAM_Mask;
    Exit;
  end;

  Result := AnalogVideo_None;
end;

class function TCaptureParameterConfig.ConvertVideoSizeInf(
  const videoSizeStr: WideString): TVideoSize;
var
  videoInf: WideString;
begin
  //video format格式“320 X 240”
  videoInf := UpperCase(videoSizeStr);

  Result.Width := 320;
  Result.Height := 240;

  if Pos(VIDEO_FORMAT_SPLIT, videoInf) <= 0 then Exit;

  //取得视频大小数据
  Result.Width := StrToInt(Trim(Copy(videoInf, 1, Pos(VIDEO_FORMAT_SPLIT, videoInf) - 1)));
  Result.Height := StrToInt(Trim(Copy(videoInf, Pos(VIDEO_FORMAT_SPLIT, videoInf) + 1, Length(videoInf))));
end;

class procedure TCaptureParameterConfig.CopyParameter(
  const sourceParameter: TCaptureParameter; var destParameter: TCaptureParameter);
begin
  destParameter.CaptureDeviceName := sourceParameter.CaptureDeviceName;
  destParameter.VideoAnalog       := sourceParameter.VideoAnalog;
  destParameter.ColorDepth        := sourceParameter.ColorDepth;
  destParameter.VideoSize         := sourceParameter.VideoSize;

  destParameter.Brightness        := sourceParameter.Brightness;
  destParameter.Contrast          := sourceParameter.Contrast;
  destParameter.Hue               := sourceParameter.Hue;
  destParameter.Saturation        := sourceParameter.Saturation;
  destParameter.Gamma             := sourceParameter.Gamma;
  destParameter.WhiteBlance       := sourceParameter.WhiteBlance;

  destParameter.EncoderName      := sourceParameter.EncoderName;
  destParameter.IsTimeLimit      := sourceParameter.IsTimeLimit;
  destParameter.LimitLength      := sourceParameter.LimitLength;
  destParameter.IsConvertGrayImg := sourceParameter.IsConvertGrayImg;
  destParameter.IsApplyImageCut  := sourceParameter.IsApplyImageCut;

  destParameter.TopRate    := sourceParameter.TopRate;
  destParameter.HeightRate := sourceParameter.HeightRate;
  destParameter.LeftRate   := sourceParameter.LeftRate;
  destParameter.WidthRate  := sourceParameter.WidthRate;

  destParameter.ParameterState := sourceParameter.ParameterState;
  destParameter.VideoShowModel := sourceParameter.VideoShowModel;
  destParameter.SnatchWay      := sourceParameter.SnatchWay;
  destParameter.IsShowState    := sourceParameter.IsShowState;
  destParameter.InputCrossbar  := sourceParameter.InputCrossbar;
  destParameter.OutputCrossbar := sourceParameter.OutputCrossbar;

  destParameter.IsAutoBrightness  := sourceParameter.IsAutoBrightness;
  destParameter.IsAutoContrast    := sourceParameter.IsAutoContrast;
  destParameter.IsAutoHue         := sourceParameter.IsAutoHue;
  destParameter.IsAutoGamma       := sourceParameter.IsAutoGamma;
  destParameter.IsAutoSaturation  := sourceParameter.IsAutoSaturation;
  destParameter.IsAutoWhiteBlance := sourceParameter.IsAutoWhiteBlance;

  destParameter.IsSoundHint := sourceParameter.IsSoundHint;
  
  destParameter.DebugFilter := sourceParameter.DebugFilter;
end;


class function TCaptureParameterConfig.GetVideoSizeStr(const width,
  height: Integer): WideString;
begin
  Result := IntToStr(width) + VIDEO_FORMAT_SPLIT + IntToStr(height);  
end;

class procedure TCaptureParameterConfig.InitCaptureParameter(
  var captureParameter: TCaptureParameter);
begin
  captureParameter.CaptureDeviceName := '';
  captureParameter.VideoAnalog := 'PAL_B';
  captureParameter.ColorDepth := 0;
  captureParameter.VideoSize := '320X240';
  captureParameter.Brightness := -1;
  captureParameter.Contrast := -1;
  captureParameter.Hue := -1;
  captureParameter.Saturation := -1;
  captureParameter.Gamma := -1;
  captureParameter.WhiteBlance := -1;
  captureParameter.EncoderName := '';
  captureParameter.IsTimeLimit := false;
  captureParameter.LimitLength := 0;
  captureParameter.IsConvertGrayImg := false;
  captureParameter.IsApplyImageCut := false;
  captureParameter.TopRate := 0;
  captureParameter.HeightRate := 0;
  captureParameter.LeftRate := 0;
  captureParameter.WidthRate := 0;
  captureParameter.ParameterState := false;
  captureParameter.VideoShowModel := smNormal;
  captureParameter.SnatchWay := swVMR;
  captureParameter.IsShowState := True;
  captureParameter.IsSoundHint := False;
  captureParameter.DebugFilter := False;
  captureParameter.ExposureWay := 0;  //0不设置，1自动，2手动
  captureParameter.ExposureValue := 0;
end;

end.
