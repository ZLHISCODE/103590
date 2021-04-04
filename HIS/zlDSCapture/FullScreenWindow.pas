{*******************************************************************************
*
*  视频的全屏显示封装
*  创建人：TJH
*  创建日前：2009-11-26
*  说明：全屏的显示使用了两种方式，如果支持VMR，则用VMR的全屏显示，
*        否则使用控件自身的全屏显示方式，但需要注意的是，DSPACK组件不支持多屏下
*        的全屏显示，始终将全屏显示在第一个显示器中，因此对DSPACK单元的显示部份
*        做了修改，以便传入指定的显示窗口句柄，然后再进行视频的全屏显示。
*
*******************************************************************************}
unit FullScreenWindow;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DirectShow9, ZLDSVideoProcess_TLB, DSPack, AppEvnts, ExtCtrls,
  StdCtrls;

type
  TfrmFullScreen = class(TForm)
    timeFocus: TTimer;
    Label1: TLabel;
    timeRefreshVideo: TTimer;
    procedure FormHide(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure timeFocusTimer(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure timeRefreshVideoTimer(Sender: TObject);
  private
    { Private declarations }

    //设置显示窗口位置
    class procedure SetFullWindowPos(fullForm: TForm; const monitorIndex: Integer);

    //不使用VMR的全屏显示处理
    class procedure FullScreenWithNotUserVmr(fullForm: TfrmFullScreen;
      videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);

    //使用VMR的全屏显示处理
    class procedure FullScreenWithUserVmr(fullForm: TfrmFullScreen;
      videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);

  public
    { Public declarations }
    FullScreenVideoWindow: TVideoWindow; //视频显示窗口
    FullScreenMonitorIndex: Integer;
    FullScreenShowModel: TShowModel;

    //全屏显示
    class function ShowFullScreen(vw: TVideoWindow;
                                  const showModel: TShowModel;
                                  const monitorIndex: Integer
                                  ): WideString;

    //退出全屏
    class function QuitFullScreen(): WideString;
    //取得是否已经进入全屏显示
    class function GetFullScreenState(): Boolean;

  end;

implementation

{$R *.dfm}

uses
  CaptureDebug, Math;

var
  frmFullScreen: TfrmFullScreen = nil; //保存全屏的显示窗口
  IsExitApp: Boolean = false;          //判断是否开始退出应用程序，如果是，则不执行窗口的隐藏(HIDE)事件

{ TfrmFullScreen }


class function TfrmFullScreen.QuitFullScreen(): WideString;
begin
  try
    Result := '';

    if not Assigned(frmFullScreen) then Exit;

    if frmFullScreen.Showing then frmFullScreen.Hide;
  except
    on e: Exception do begin
      Result := e.Message;
      
      TDebug.DebugMsg('TfrmFullScreen', 'QuitFullScreen', e.Message);
    end;
  end;
end;


class function TfrmFullScreen.ShowFullScreen(
  vw: TVideoWindow;
  const showModel: TShowModel;
  const monitorIndex: Integer): WideString;
var
  vmrWindowLessCtrl9: IVMRWindowlessControl9;
  hr: HRESULT;
begin
  try
    Result := '';

    if not Assigned(frmFullScreen) then
      frmFullScreen := TfrmFullScreen.Create(Application{nil});

    frmFullScreen.OnClick := vw.OnClick;
    frmFullScreen.OnDblClick := vw.OnDblClick;
    frmFullScreen.OnKeyDown := vw.OnKeyDown;
    frmFullScreen.OnKeyPress := vw.OnKeyPress;
    frmFullScreen.OnKeyUp := vw.OnKeyUp;
    frmFullScreen.OnMouseDown := vw.OnMouseDown;
    frmFullScreen.OnMouseMove := vw.OnMouseMove;
    frmFullScreen.OnMouseUp := vw.OnMouseUp;
    frmFullScreen.OnPaint := vw.OnPaint;

    frmFullScreen.Color := vw.Color;
    frmFullScreen.FullScreenVideoWindow := vw;
    frmFullScreen.FullScreenMonitorIndex := monitorIndex;
    frmFullScreen.FullScreenShowModel := showModel;

    frmFullScreen.Show;

    SetFullWindowPos(frmFullScreen, monitorIndex);

    hr := (vw as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtrl9);
    if Failed(hr) then begin
      //不使用VMR的全屏处理
      FullScreenWithNotUserVmr(frmFullScreen, vw, showModel, monitorIndex);
    end else begin
      //使用VMR显示全屏
      vmrWindowLessCtrl9 := nil;
      FullScreenWithUserVmr(frmFullScreen, vw, showModel, monitorIndex);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
      if Assigned(frmFullScreen) then FreeAndNil(frmFullScreen);
    end;
  end;
end;


procedure TfrmFullScreen.FormHide(Sender: TObject);
var
  vmrWindowLessCtr9: IVMRWindowlessControl9;
  hr: HRESULT;
  videoRect, localRect: TRect;
  localWidth, localHeight, arWidth, arHeight: Integer;
begin
  try
    if IsExitApp then Exit;
    if (FullScreenVideoWindow = nil) then Exit;
    
    hr := (FullScreenVideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtr9);
    if Failed(hr) then begin
      FullScreenVideoWindow.FullScreen := False;
      Exit;
    end;

    try
      if (vmrWindowLessCtr9 = nil) then Exit;
      
      vmrWindowLessCtr9.GetNativeVideoSize(localWidth, localHeight, arWidth, arHeight);

      videoRect := Rect(0, 0, FullScreenVideoWindow.Width, FullScreenVideoWindow.Height);
      localRect := Rect(0, 0, localWidth, localHeight);

      vmrWindowLessCtr9.SetVideoClippingWindow(FullScreenVideoWindow.Handle);
      vmrWindowLessCtr9.SetVideoPosition(@localRect, @videoRect);

      FullScreenVideoWindow.VMROptions.KeepAspectRatio := False;

      FullScreenVideoWindow := nil;
    finally
      vmrWindowLessCtr9 := nil;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TfrmFullScreen', 'FormHide', e.Message);
    end;
  end;
end;

procedure TfrmFullScreen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frmFullScreen := nil;
  Action := caFree;
end;


class function TfrmFullScreen.GetFullScreenState(): Boolean;
begin
  try
    Result := False;

    if not Assigned(frmFullScreen) then Exit;

    Result := frmFullScreen.Showing;
  except
    Result := False;
  end;
end;

class procedure TfrmFullScreen.SetFullWindowPos(fullForm: TForm; const monitorIndex: Integer);
begin
  fullForm.SetBounds(Screen.Monitors[monitorIndex].Left,
                     Screen.Monitors[monitorIndex].Top,
                     Screen.Monitors[monitorIndex].Width,
                     Screen.Monitors[monitorIndex].Height);
end;

class procedure TfrmFullScreen.FullScreenWithNotUserVmr(fullForm: TfrmFullScreen;
  videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);
begin
  videoWindow.IsStreach := ShowModel = smStretch;
  videoWindow.MonitorIndex := MonitorIndex;
  videoWindow.FullHandle := frmFullScreen.Handle;
  videoWindow.FullScreen := True;
end;

class procedure TfrmFullScreen.FullScreenWithUserVmr(fullForm: TfrmFullScreen;
  videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);
var
  hr: HRESULT;
  screenRect, videoRect: TRect;
  vmrWindowLessCtr9: IVMRWindowlessControl9;
  localWidth, localHeight, arWidth, arHeight: Integer;
begin
  try
    hr := (videoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtr9);
    if Failed(hr) then begin
      fullForm.Label1.Visible := True;
      Exit;
    end;

    fullForm.Label1.Visible := False;
    
    case showModel of
      smNormal, smFit, smAutoFitCut, smWindAutoFit: begin
        VideoWindow.VMROptions.KeepAspectRatio := True;
      end;
      smStretch: begin
        VideoWindow.VMROptions.KeepAspectRatio := False;
      end;
    end;

    vmrWindowLessCtr9.GetNativeVideoSize(localWidth, localHeight, arWidth, arHeight);

    screenRect := Rect(0, 0, Screen.Monitors[MonitorIndex].Width, Screen.Monitors[MonitorIndex].Height);
    videoRect := Rect(0, 0, localWidth, localHeight);

    vmrWindowLessCtr9.SetVideoClippingWindow(fullForm.Handle);
    vmrWindowLessCtr9.SetVideoPosition(@videoRect, @screenRect);
  finally
    vmrWindowLessCtr9 := nil;
  end;
end;

procedure TfrmFullScreen.timeFocusTimer(Sender: TObject);
begin
  //当调用该组件的程序中使用全屏时，如果在全屏的视频窗口中点击了鼠标按键，
  //则该窗口将失去焦点，因此需要使用计数器获取失去的焦点，
  //用Deactivate事件或者截获消息，都不能有效的重获该窗口的焦点
  if Self.Visible then begin
    SetFocus;

    if Focused then begin
      timeFocus.Enabled := False;
    end;  
  end;
end;

procedure TfrmFullScreen.FormDeactivate(Sender: TObject);
begin
  timeFocus.Enabled := True;
end;

procedure TfrmFullScreen.timeRefreshVideoTimer(Sender: TObject);
begin
  if Self.Visible then begin
    //使用计数器刷新VMR全屏窗口是因为在任务栏最小化全屏后，在显示时，视频只显示了一小块区域
    FullScreenWithUserVmr(Self,
                          FullScreenVideoWindow,
                          FullScreenShowModel,
                          FullScreenMonitorIndex);
  end;
end;

initialization

finalization
  if Assigned(frmFullScreen) then begin
    IsExitApp := True; //如果这里不设置程序的退出变量为TRUE，当使用该组件的应用程序退出时，在该窗口的HIDE事件中将会引发异常。
    FreeAndNil(frmFullScreen);
  end;  

end.
