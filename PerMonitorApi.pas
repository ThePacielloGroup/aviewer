unit PerMonitorApi;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, MultiMon, Vcl.ExtCtrls;


const
   Process_DPI_Unaware = 0;
   Process_System_DPI_Aware = 1;
   Process_Per_Monitor_DPI_Aware = 2;
   E_ACCESSDENIED = $80070005;
   WM_DPICHANGED = 736;
type
  MONITOR_DPI_TYPE = (
    MDT_EFFECTIVE_DPI,
    MDT_ANGULAR_DPI,
    MDT_RAW_DPI
  );
  TGetProcessDPIAwarenessProc = function(const hprocess: THandle; var ProcessDpiAwareness: LongInt): HRESULT; stdcall;
   TSetProcessDPIAwarenessProc = function(const ProcessDpiAwareness: LongInt): HRESULT; stdcall;
  function GetWindowMonitorDPI(Handle: HWND): Integer;
  function SystemCanSupportPerMonitorDpi(AutoEnable: Boolean): Boolean; // New Windows 8.1 dpi awareness available?

  function SystemCanSupportOldDpiAwareness(AutoEnable: Boolean): Boolean; // Windows Vista/ Windows 7 Global System DPI functional level.
  function GetDpiForMonitor(hmonitor: HMONITOR; dpiType: MONITOR_DPI_TYPE; out dpiX: UINT; out dpiY: UINT): HRESULT; external 'shcore.dll';

  procedure GetWindowScale(Handle: HWND; Dx, Dy: double; var SX, SY: Double);
  procedure GetDCap(Handle: HWND; var SX, SY: double);
var
   _RequestedLevelOfAwareness:LongInt;
   _ProcessDpiAwarenessValue:LongInt;

implementation




function _GetProcessDpiAwareness(AutoEnable: Boolean): LongInt;
var
   hprocess: THandle;
   HRESULT: DWORD;
   BAwareness: Integer;
   GetProcessDPIAwareness: TGetProcessDPIAwarenessProc;
   LibHandle: THandle;
   PID: DWORD;

   function ManifestOverride: Boolean;
   var
      HRESULT: DWORD;
      SetProcessDPIAwareness: TSetProcessDPIAwarenessProc;
   begin
      Result := False;
      SetProcessDPIAwareness := TSetProcessDPIAwarenessProc(GetProcAddress(LibHandle, 'SetProcessDpiAwareness'));
      if Assigned(SetProcessDPIAwareness) and (_RequestedLevelOfAwareness>=0) then
      begin
         HRESULT := SetProcessDPIAwareness(_RequestedLevelOfAwareness );
         Result := (HRESULT = 0) or (HRESULT = E_ACCESSDENIED)
      end
   end;

begin
   Result := _ProcessDpiAwarenessValue;
   if (Result = -1) then
   begin
      BAwareness := 3;
      LibHandle := LoadLibrary('shcore.dll');
      if LibHandle <> 0 then
      begin
         if (not AutoEnable) or ManifestOverride then
         begin
            GetProcessDPIAwareness := TGetProcessDPIAwarenessProc(GetProcAddress(LibHandle, 'GetProcessDpiAwareness'));
            if Assigned(GetProcessDPIAwareness) then
            begin
               PID := WinApi.Windows.GetCurrentProcessId;
               hprocess := OpenProcess(PROCESS_ALL_ACCESS, False, PID);
               if hprocess > 0 then
               begin
                  HRESULT := GetProcessDPIAwareness(hprocess, BAwareness);
                  if HRESULT = 0 then
                     Result := BAwareness;
               end;
            end;
         end;
      end;
   end;
end;

function SystemCanSupportPerMonitorDpi(AutoEnable: Boolean): Boolean;
begin
   if AutoEnable then
   begin
    _RequestedLevelOfAwareness := Process_Per_Monitor_DPI_Aware;
    _ProcessDpiAwarenessValue := -1;
   end;
   Result := _GetProcessDpiAwareness(AutoEnable) = Process_Per_Monitor_DPI_Aware;
end;

function SystemCanSupportOldDpiAwareness(AutoEnable: Boolean): Boolean;
begin
   if AutoEnable then
   begin
     _RequestedLevelOfAwareness := Process_Per_Monitor_DPI_Aware;
     _ProcessDpiAwarenessValue := -1;
   end;

   Result := _GetProcessDpiAwareness(AutoEnable) = Process_System_DPI_Aware;
end;

function GetWindowMonitorDPI(Handle: HWND): Integer;
var
  HorzDPI: UINT;
  VertDPI: UINT;
  Monitor: HMONITOR;
  DC: HDC;
  function IsWin81: boolean;
  var
      VI: TOSVersionInfo;
  begin
    Result := False;
    FillChar(VI, SizeOf(VI), 0);
    VI.dwOSVersionInfoSize := SizeOf(VI);

    if GetVersionEx(VI) then
    begin

        if (VI.dwMajorVersion >= 6) and (VI.dwMinorVersion >= 3) then
            Result := True;
    end;
end;
begin
  if IsWin81 then
  begin
    Result := 96;
    Monitor := MonitorFromWindow(Handle, MONITOR_DEFAULTTONULL);
    if Monitor <> 0 then
    begin
      if GetDpiForMonitor(Monitor, MDT_EFFECTIVE_DPI, HorzDPI, VertDPI) = S_OK then
        Result := VertDPI;
    end;
  end
  else
  begin
    DC := GetDC(0);
    Result := GetDeviceCaps(DC, LOGPIXELSY);
    ReleaseDC(0, DC);
  end;
end;

procedure GetWindowScale(Handle: HWND; Dx, Dy: double; var SX, SY: Double);
var
  i: Integer;
  dc: HDC;
  monEx: TMonitorInfoEx;
  hm: HMonitor;
begin
    FillChar(monEx, SizeOf(TMonitorInfoEx), #0);
    for i := 0 to Screen.MonitorCount - 1 do
    begin


      GetMonitorInfo(Screen.Monitors[i].Handle, @monEx);
      hm := MonitorFromWindow(Handle, MONITOR_DEFAULTTONEAREST);
      if hm = Screen.Monitors[i].Handle then
      begin
        dc := CreateDC('DISPLAY', monEx.szDevice, nil, nil);
        try
          SX := GetDeviceCaps(dc, LOGPIXELSX) / Dx;
          SY := GetDeviceCaps(dc, LOGPIXELSY) / Dy;
        finally
          DeleteDC(dc);
        end;
      end;

    end;
end;

procedure GetDCap(Handle: HWND; var SX, SY: double);
var
  i: Integer;
  dc: HDC;
  monEx: TMonitorInfoEx;
  hm: HMonitor;
begin
    FillChar(monEx, SizeOf(TMonitorInfoEx), #0);
    for i := 0 to Screen.MonitorCount - 1 do
    begin


      GetMonitorInfo(Screen.Monitors[i].Handle, @monEx);
      hm := MonitorFromWindow(Handle, MONITOR_DEFAULTTONEAREST);
      if hm = Screen.Monitors[i].Handle then
      begin
        dc := CreateDC('DISPLAY', monEx.szDevice, nil, nil);
        try
          SX := GetDeviceCaps(dc, LOGPIXELSX);
          SY := GetDeviceCaps(dc, LOGPIXELSY);
        finally
          DeleteDC(dc);
        end;
      end;

    end;


end;

initialization
   _ProcessDpiAwarenessValue := -1;
   _RequestedLevelOfAwareness := -1;

end.
