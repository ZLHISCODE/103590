unit DSCapParameterConfigObj;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, ZLDSVideoProcess_TLB, StdVcl, SysUtils;

type
  TDSCapParameterEnum = class(TAutoObject, IDSParameterEnum)
  private
  public
    //ȡ���豸����
    function GetDeviceCount(var deviceCount: SYSINT): WideString; safecall;
    //ȡ���豸����
    function GetDeviceName(deviceIndex: SYSINT; var deviceName: WideString): WideString; safecall;
    //ȡ�ñ���������
    function GetEncoderCount(var encoderCount: SYSINT): WideString; safecall;
    //ȡ�ñ���������
    function GetEncoderName(encoderIndex: SYSINT; var encoderName: WideString): WideString; safecall;
    //ȡ����Ƶ��ɫ���
    function GetVideoColorDepth(colorDepthIndex: SYSINT; var colorDepth: SYSINT): WideString; safecall;
    //ȡ����ɫ�������
    function GetVideoColorDepthCount(var colorDepthCount: SYSINT): WideString; safecall;
    //ȡ����Ƶ��ʽ����
    function GetVideoAnalogCount(var analogCount: SYSINT): WideString; safecall;
    //ȡ����Ƶ��ʽ����
    function GetVideoAnalogName(analogIndex: SYSINT; var analogName: WideString): WideString; safecall;
    //ȡ���豸�������������ֵ
    function GetVideoQualityMaxValue(const deviceName: WideString; qualityType: TQualityType; var maxValue: SYSINT): WideString; safecall;
    //ȡ�÷ֱ�������
    function GetVideoSizeCount(var sizeCount: SYSINT): WideString; safecall;
    //ȡ�÷ֱ�������
    function GetVideoSizeName(sizeIndex: SYSINT; var sizeName: WideString): WideString; safecall;
    //����Ƿ�ΪVFW�豸
    function CheckIsVfwDevice(const deviceName: WideString): WordBool; safecall;
    //����Ƿ�֧��VMRģʽ
    function CheckIsSupportVmr: WordBool; safecall;
    //��Ƶ��С��ʽת��
    function VideoSizeConvert(const videoSize: WideString): TVideoSize; safecall;
    //�ж��Ƿ�֧����Ƶ��������
    function GetIsSupportQuailtiCfg(const deviceName: WideString): WordBool; safecall;
  end;

implementation

uses ComServ, DirectShow9, DirectShow9Ex, DSUtil, CaptureDebug, VideoProcessDefine;

function TDSCapParameterEnum.GetDeviceCount(
  var deviceCount: SYSINT): WideString;
var
  captureDeviceEnum: TSysDevEnum;
begin
  try
    Result := '';
    deviceCount := 0;

    //ö�ٲɼ��豸.
    captureDeviceEnum := TSysDevEnum.Create(CLSID_VideoInputDeviceCategory);
    if not Assigned(captureDeviceEnum) then Exit;

    deviceCount := captureDeviceEnum.CountFilters;
    FreeAndNil(captureDeviceEnum);
  except
    on e: Exception do begin
      deviceCount := 0;
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetDeviceName(deviceIndex: SYSINT;
  var deviceName: WideString): WideString;
var
  captureDeviceEnum: TSysDevEnum;
begin
  try
    Result := '';
    deviceName := '';
    //ö�ٲɼ��豸.
    captureDeviceEnum := TSysDevEnum.Create(CLSID_VideoInputDeviceCategory);
    if not Assigned(captureDeviceEnum) then Exit;

    try
      //���豸��ӵ��б���
      deviceName := captureDeviceEnum.Filters[deviceIndex].FriendlyName;
    finally
      FreeAndNil(captureDeviceEnum);
    end;
  except
    on e: Exception do begin
      deviceName := '';
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetEncoderCount(
  var encoderCount: SYSINT): WideString;
var
  videoEncoderEnum: TSysDevEnum;
begin
  try
    Result := '';
    encoderCount := 0;

    //ö�ٱ����豸
    videoEncoderEnum  := TSysDevEnum.Create(CLSID_VideoCompressorCategory);
    if not Assigned(videoEncoderEnum) then Exit;

    encoderCount := videoEncoderEnum.CountFilters;
    FreeAndNil(videoEncoderEnum);
  except
    on e: Exception do begin
      encoderCount := 0;
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetEncoderName(encoderIndex: SYSINT;
  var encoderName: WideString): WideString;
var
  videoEncoderEnum: TSysDevEnum;
begin
  try
    Result := '';

    encoderName := '';

    //ö�ٱ����豸
    videoEncoderEnum  := TSysDevEnum.Create(CLSID_VideoCompressorCategory);
    if not Assigned(videoEncoderEnum) then Exit;

    try
      encoderName := videoEncoderEnum.Filters[encoderIndex].FriendlyName;
    finally
      FreeAndNil(videoEncoderEnum);
    end;
  except
    on e: Exception do begin
      encoderName := '';
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoAnalogCount(
  var analogCount: SYSINT): WideString;
//const
//  CurrentUseVideoModeCount: Integer = 23;
begin
  try
    Result := '';

    //����Ĭ�ϵ���Ƶ��ʽ����
    analogCount := Length(SysVideoAnalog); //CurrentUseVideoModeCount;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoAnalogName(analogIndex: SYSINT;
  var analogName: WideString): WideString;
begin
  try
    Result := '';

    if (analogIndex < 0) or (analogIndex >= Length(SysVideoAnalog)) then begin
      analogName := '';
      Exit;
    end;

    analogName := SysVideoAnalog[analogIndex];

    {
    //����Ĭ�ϵ���Ƶ��ʽ
    case analogIndex of
      0: analogName := 'PAL_B';
      1: analogName := 'PAL_D';
      2: analogName := 'PAL_G';
      3: analogName := 'PAL_H';
      4: analogName := 'PAL_I';
      5: analogName := 'PAL_M';
      6: analogName := 'PAL_N';
      7: analogName := 'PAL_60';
      8: analogName := 'PAL_Mask';
      9: analogName := 'NTSC_M';
      10: analogName := 'NTSC_M_J';
      11: analogName := 'NTSC_433';
      12: analogName := 'NTSC_Mask';
      13: analogName := 'SECAM_B';
      14: analogName := 'SECAM_D';
      15: analogName := 'SECAM_G';
      16: analogName := 'SECAM_H';
      17: analogName := 'SECAM_K';
      18: analogName := 'SECAM_Kl';
      19: analogName := 'SECAM_L';
      20: analogName := 'SECAM_L1';
      21: analogName := 'SECAM_Mask';
      22: analogName := 'None';
      else analogName := '';
    end;}

  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoColorDepth(colorDepthIndex: SYSINT;
  var colorDepth: SYSINT): WideString;
begin
  try
    Result := '';

    //���ع̶�����ɫ���
    case colorDepthIndex of
      0: colorDepth := 8;
      1: colorDepth := 12;
      2: colorDepth := 16;
      3: colorDepth := 24;
      4: colorDepth := 32;
      else colorDepth := 0;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoColorDepthCount(
  var colorDepthCount: SYSINT): WideString;
const
  CurrentUseColorDepthCount: Integer = 5;
begin
  try
    Result := '';

    //ȡ����ɫ�������
    colorDepthCount := CurrentUseColorDepthCount;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoQualityMaxValue(
  const deviceName: WideString; qualityType: TQualityType;
  var maxValue: SYSINT): WideString;
var
  captureFilter: IBaseFilter;
  amVideoProcAmp: IAMVideoProcAmp;
  hr: HRESULT;

  //ȡ���������������Ϣ
  function GetMaxValue(curAmVideoProcAmp: IAMVideoProcAmp;
                   PropertyTag : TVideoProcAmpProperty): Integer;
  var
    curHr: HRESULT;
    iMinValue, iMaxValue, iStep, iDefault: Integer;
    iFlags : TVideoProcAmpFlags;
  begin

    //ȡ����Ƶ�������õķ�Χ
    curHr := curAmVideoProcAmp.GetRange(PropertyTag, iMinValue, iMaxValue, iStep, iDefault, iFlags);
    if not Succeeded(curHr) then begin
      Result := 0;
      Exit;
    end;

    Result := iMaxValue;
  end;

begin
  try
    Result := '';
    maxValue := 0; 

    //����vfw���豸,�򲻽��ж�ȡ
    if TDS9Ex.IsVfwDevice(deviceName) then Exit;

    hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, deviceName, captureFilter);
    if Failed(hr) then begin
      Result := '�����ɼ��豸�ӿ�ʱʧ�ܡ�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    try
      hr := captureFilter.QueryInterface(IID_IAMVideoProcAmp, amVideoProcAmp);
      if Failed(hr) then begin
        Result := '������Ƶ�������ýӿ�ʱʧ�ܣ��豸��֧�ָ����á�[�������:' + IntToStr(hr) + ']';
        Exit;
      end;

      try
        //ȡ����Ƶ���������ֵ
        case qualityType of
          qtBrightness: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_Brightness);
          qtContrast: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_Contrast);
          qtHue: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_Hue);
          qtSaturation: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_Saturation);
          qtGamma: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_Gamma);
          qtWhiteBlance: maxValue := GetMaxValue(amVideoProcAmp, VideoProcAmp_WhiteBalance);
          else maxValue := 0;
        end;
      finally
        amVideoProcAmp := nil;
      end;

    finally
      captureFilter := nil;
    end;
  except
    on e: Exception do begin
      maxValue := 0;
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoSizeCount(
  var sizeCount: SYSINT): WideString;
//const
//  CurrentUseVideoSizeCount: Integer = 11;
begin
  try
    Result := '';

    //ȡ��Ĭ�Ϸֱ��ʵ�����
    sizeCount := Length(SysVideoSize);// CurrentUseVideoSizeCount;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.GetVideoSizeName(sizeIndex: SYSINT;
  var sizeName: WideString): WideString;
begin
  try
    Result := '';

    if (sizeIndex < 0) or (sizeIndex >= Length(SysVideoSize)) then begin
      sizeName := '';
      Exit;
    end;

    sizeName := SysVideoSize[sizeIndex];
                     
    {//����Ĭ�Ϸֱ������
    case sizeIndex of
      0: sizeName := '160X120';
      1: sizeName := '176X144';
      2: sizeName := '240X180';
      3: sizeName := '320X240';
      4: sizeName := '352X288';
      5: sizeName := '512X380';
      6: sizeName := '640X480';
      7: sizeName := '704X576';
      8: sizeName := '720X576';
      9: sizeName := '768X576';
      10: sizeName := '800X600';
      else sizeName := '';
    end;}
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSCapParameterEnum.CheckIsVfwDevice(
  const deviceName: WideString): WordBool;
begin
  Result := TDS9Ex.IsVfwDevice(deviceName);
end;

function TDSCapParameterEnum.CheckIsSupportVmr: WordBool;
var
  AFilter: IBaseFilter;
  CW: Word;
begin
  CW := Get8087CW;
  try
    result := (CoCreateInstance(CLSID_VideoMixingRenderer9, nil, CLSCTX_INPROC, IID_IBaseFilter ,AFilter) = S_OK);
  finally
    Set8087CW(CW);
    AFilter := nil;
  end;
end;

function TDSCapParameterEnum.VideoSizeConvert(
  const videoSize: WideString): TVideoSize;
begin
  Result := TCaptureParameterConfig.ConvertVideoSizeInf(videoSize);
end;

function TDSCapParameterEnum.GetIsSupportQuailtiCfg(
  const deviceName: WideString): WordBool;
var
  captureFilter: IBaseFilter;
  amVideoProcAmp: IAMVideoProcAmp;
  hr: HRESULT;

begin
    Result := False;

    //����vfw���豸,�򲻽��ж�ȡ
    if TDS9Ex.IsVfwDevice(deviceName) then Exit;

    hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, deviceName, captureFilter);
    if Failed(hr) then Exit;

    try
      hr := captureFilter.QueryInterface(IID_IAMVideoProcAmp, amVideoProcAmp);
      if Failed(hr) then Exit;

      Result := True;
    finally
      captureFilter := nil;
    end;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TDSCapParameterEnum, Class_DSCapParameterEnum,
    ciMultiInstance, tmApartment);
end.
