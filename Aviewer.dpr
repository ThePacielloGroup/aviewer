program Aviewer;
(*
aViewer is an accessibility API object inspection tool.

Copyright (C) 2014 The Paciello Group

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
*)
{$R 'bmp.res' 'bmp.rc'}
{$R 'about.res' 'about.rc'}
uses
  Forms,
  Windows,
  SysUtils,
  Messages,
  Dialogs,
  Registry,
  System.IOUtils,
  FocusRectWnd in 'FocusRectWnd.pas' {WndFocusRect},
  frmMSAAV in 'frmMSAAV.pas' {wndMSAAV},
  TipWnd in 'TipWnd.pas' {frmTipWnd},
  frmSet in 'frmSet.pas' {WndSet},
  Thread in 'Thread.pas';

{$R *.res}
var
    wnd, AppWnd: Thandle;

const
  AppUniqueName:string = 'Global\msaav-KGZ6PhtR3P37';

function IsPrevAppExist(Name:string):Boolean;
begin
  result := false;
  CreateMutex(nil,true,PChar(Name));
  if GetLastError = ERROR_ALREADY_EXISTS then
    result := true;
end;

procedure CopyDataToOld;
var
    i: Integer;
    SendData: TCOPYDATASTRUCT;
    strBuf: string;
begin
    if ParamCount > 0 then
    begin
        //strBuf := '"' + IntToStr(ParamCount) + '",';
        for i := 1 to ParamCount do
        begin
            strBuf := StrBuf + ParamStr(i);
            if i <> ParamCount then
                strBuf := strBuf + ',';
        end;
        SendData.dwData := $00000325;
        SendData.cbData := length(strBuf) * SizeOf(Char);
        SendData.lpData := StrAlloc(SendData.cbData);
        try
          StrCopy(SendData.lpData, PChar(strBuf));
          SendMessage(Wnd,WM_COPYDATA,0,LPARAM(@SendData));
        finally

          SysUtils.StrDispose(Pchar(SendData.lpData));
        end;

    end;
end;

procedure RegFBE;
var
	Reg: TRegistry;
  sFN: string;
  cType: Cardinal;
  lFlag: LongBool;
begin
  lFlag := GetBinaryType(PChar(application.ExeName), cType);
  if lFlag then
  begin
    if cType = SCS_64BIT_BINARY then
    	sFN := 'aViewer64bit.exe'
    else
    	sFN := 'aViewer32bit.exe';

  end
  else
    sFN := 'aViewer32bit.exe';
  Reg := TRegistry.Create(KEY_ALL_ACCESS );
  try
  	Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\', false) then
    begin

      if (not Reg.ValueExists(sFN)) then
      	Reg.WriteInteger(sFN, $00002af8);

      Reg.CloseKey;
    end;
    {if Reg.OpenKey('SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\', False) then
    begin
			if (not Reg.ValueExists(sFN)) then
      	Reg.WriteInteger(sFN, $00002af8);
      Reg.CloseKey;
    end;}

  finally
  	Reg.Free;
  end;
end;



begin
    RegFBE;
    if IsPrevAppExist(AppUniqueName) then
    begin
        Wnd := FindWindow('TwndMSAAV', 'Accessibility Viewer');
        if (ParamCount = 0) then
        begin
            if Wnd <> 0 then
            begin
                AppWnd := GetWindowLong(Wnd, GWL_HWNDPARENT);
                if AppWnd <> 0 then Wnd := AppWnd;
                if IsIconic(Wnd) then
                    OpenIcon(Wnd)
                else
                    SetForegroundWindow(Wnd);
            end;
        end
        else
        begin
            if Wnd <> 0 then
            begin
                CopyDataToOld;
                AppWnd := GetWindowLong(Wnd, GWL_HWNDPARENT);
                if AppWnd <> 0 then Wnd := AppWnd;
                if IsIconic(Wnd) then
                    OpenIcon(Wnd)
                else
                    SetForegroundWindow(Wnd);
            end;
        end;
        Application.Terminate;
    end
    else
    begin

    	Application.Initialize;
      Application.MainFormOnTaskbar := True;
      Application.CreateForm(TwndMSAAV, wndMSAAV);
  Application.Run;
    end;
end.
