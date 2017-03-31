unit frmSet;
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
interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls,
  CheckLst, AccCTRLs, commctrl, iniFiles, ComObj, ShlObj, ShellAPI, Math, StrUtils;

type
  TWndSet = class(TForm)
    gbSelProp: TAccGroupBox;
    chkList: TAccCheckList;
    gbFont: TAccGroupBox;
    btnFont: TAccButton;
    gbRegDll: TAccGroupBox;
    btnReg: TAccButton;
    btnUnreg: TAccButton;
    btnOK: TAccButton;
    btnCancel: TAccButton;
    cbExTip: TTransCheckBox;
    FontDialog1: TFontDialog;
    Memo1: TAccMemo;
    procedure FormCreate(Sender: TObject);
    procedure btnRegClick(Sender: TObject);
    procedure btnUnregClick(Sender: TObject);
    procedure btnFontClick(Sender: TObject);
  private
    { Private declare }

    Transpath, APPDir, SPath, DllPath, M1: string;

    procedure WMDPIChanged(var Message: TMessage); message WM_DPICHANGED;
  public
    { Public declare }
    DefFont: integer;
    ScaleX, ScaleY, DefX, DefY: double;
    procedure Load_Str(Path: string);
    procedure SizeChange;
  end;

var
  WndSet: TWndSet;

implementation

{$R *.dfm}

function DoubleToInt(d: double): integer;
begin
  SetRoundMode(rmUP);
  Result := Trunc(SimpleRoundTo(d));
end;

function IsWinVista: boolean;
var
    VI: TOSVersionInfo;
begin
    Result := False;
    FillChar(VI, SizeOf(VI), 0);
    VI.dwOSVersionInfoSize := SizeOf(VI);
    if GetVersionEx(VI) then
    begin
        if (VI.dwMajorVersion >= 6) then
            Result := True;
    end;
end;
function IsLimited: boolean;
var
    hProcess, hToken : THandle;
    pt: TTokenElevationType;
    dwLength: DWORD;
begin

    Result := True;
    if not IsWinVista then
    begin
        Result := False;
        Exit;
    end;
    hProcess := GetCurrentProcess();
    if OpenProcessToken(hProcess, TOKEN_QUERY {or TOKEN_QUERY_SOURCE}, hToken) then
    begin
        if GetTokenInformation(hToken, Windows.TTokenInformationClass(TokenElevationType), @pt, sizeOf(@pt), dwLength) then
        begin
            if pt <> TokenElevationTypeLimited then
                Result := False;
            CloseHandle(hToken);

        end;
    end;
end;
function Execute(FileName, Param: string; RunAdmin: boolean = false): cardinal;
var
    op: string;
begin
    if RunAdmin then
        op := 'runas'
    else
        op := 'open';
    result := ShellExecute(0, Pchar(op), Pchar(FileName), Pchar(Param), nil, SW_SHOWNORMAL);
end;

function GetMyDocPath: string;
var
  IIDList: PItemIDList;
  buffer: array [0..MAX_PATH - 1] of char;
begin
  IIDList := nil;

  OleCheck(SHGetSpecialFolderLocation(Application.Handle,
    CSIDL_PERSONAL, IIDList));
  if not SHGetPathFromIDList(IIDList, buffer) then
  begin
    raise Exception.Create('A virtual diractory cannot be acquired.');
  end;
  Result := StrPas(Buffer);
end;

procedure TWndSet.SizeChange;
var
    SZ, SZ2, SZ3, SZ4: TSize;
    iMax, iMax2: integer;

    procedure GetStrSize(Cap: string; var FSZ: TSize);
    begin

      FSZ.cx := DoubleToInt(Canvas.TextWidth(Cap));
      FSZ.cy := DoubleToInt(Canvas.TextHeight(Cap));

    end;
begin
    Font.Size := DoubleToInt(DefFont * ScaleX);
    M1 := Font.Name + ',' + IntToStr(DefFont);

    Memo1.Lines[0] := '';
    Memo1.Lines[0] := M1;
    GetStrSize(gbSelProp.Caption, SZ);
    GetStrSize(gbFont.Caption, SZ2);
    GetStrSize(btnFont.Caption, SZ3);
    GetStrSize(M1, SZ4);
        //ChkList.Top := sz.Height + 5;

    if SZ.Width> ChkList.Width then
    begin
        gbSelProp.Width := sz.Width + 20;
    end;
    ChkList.Top := SZ.Height;
    ChkList.Width := gbSelProp.Width - 30;
    gbSelProp.Height := ChkList.Height + SZ.Height + 10;
    gbFont.Left := gbSelProp.Left + gbSelProp.Width + 6;

    btnFont.Top := SZ2.Height+5;
    btnFont.Width := SZ3.Width + 5;
    btnFont.Height := SZ3.Height + 5;
    Memo1.Top := btnFont.Top + BtnFont.Height + 6;
    memo1.Height := SZ4.Height * 3;
    memo1.Width := SZ4.width *2;
    iMax := Max(Memo1.Width, SZ3.Width);
    if SZ2.Width> iMax then
    begin
        gbFont.Width := sz2.Width + 32;
    end
    else
    begin
        gbFont.Width := iMax + 32;
    end;
    gbFont.Height := 50 + btnFont.Height + Memo1.Height + (sz.Height div 2);
    gbRegDll.Left := gbSelProp.Left + gbSelProp.Width + 6;
    gbRegDll.Top := gbFont.Top + gbFont.Height + 6;
    GetStrSize(gbRegDll.Caption, SZ);
    GetStrSize(btnReg.Caption, SZ2);
    GetStrSize(btnunReg.Caption, SZ3);

    iMax := Max(SZ2.Width, SZ3.Width);

    if SZ.Width > iMax then
    begin
        gbRegDll.Width := sz.Width + 32;
    end
    else
        gbRegDll.Width := iMax + 32;
    btnReg.Top := SZ.Height+5;
    btnReg.Width := iMax + 16 + 5;
    btnReg.Height := SZ2.Height + 5;
    btnunReg.Top := btnReg.Top + btnReg.Height + 5;
    btnUnReg.Height := SZ3.Height + 5;
    btnunReg.Width := iMax + 16 + 5;
    gbRegDll.Height := 50 + btnUnReg.Height + btnReg.Height;

    GetStrSize(cbExTip.Caption, SZ);


    iMax := Max(gbRegDll.Width, gbFont.Width);
    iMax := Max(iMax, cbExtip.Width);
    iMax2 := Max(gbSelProp.Height, gbFont.Height + gbRegDll.Height {+ btnReg.Height + btnUnreg.Height}{ + cbExtip.Height});
    width := gbSelProp.Width + iMax + 40;
    Height := iMax2 + 50;


    GetStrSize(btnOK.Caption, SZ);
    GetStrSize(btnCancel.Caption, SZ2);
    iMax := Max(SZ2.Width, SZ.Width);

    btnOK.Height := SZ.Height + 5;
    btnOK.Width := iMax + 10;


    btnCancel.Height := SZ2.Height + 5;
    btnCancel.Width := iMax + 10;
    btnOK.Left := width - 20 - iMax;
    btnCancel.Left := width - 20 - iMax;
    btnOK.Top := ClientHeight - (btnOK.Height + btnCancel.Height + 10);
    btnCancel.Top := btnOK.Top + btnOK.Height + 5;

end;

procedure TWndSet.btnFontClick(Sender: TObject);

begin

    FontDialog1.Font.Name := Font.Name;
    FontDialog1.Font.Charset := Font.Charset;
    FontDialog1.Font.Size := DefFont;
    if FontDialog1.Execute then
    begin
        Font := FontDialog1.Font;
        DefFont := Font.Size;
        SizeChange;
    end;

end;

function IsWOW64: Boolean;
{$IFDEF MSWINDOWS}
{$IFDEF WIN32}
var
  Flg: Bool;
{$ENDIF}
{$ENDIF}
begin
  result := False;
  {$IFDEF MSWINDOWS}
  {$IFDEF WIN32}
  if IsWow64Process(GetCurrentProcess, Flg) then
  	result := Flg;
  {$ENDIF}
  {$ENDIF}
end;

procedure TWndSet.btnRegClick(Sender: TObject);
var
    fn, pm:string;
    p: PWideChar;
const
	FOLDERID_SystemX86 : KNOWNFOLDERID ='{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}';
begin
    fn := 'RegSvr32';
    if IsWOW64 then
    begin
      if SHGetKnownFolderPath(FOLDERID_SystemX86, 0, 0, p) = S_OK then
      	fn := IncludeTrailingPathDelimiter(p) + 'RegSvr32';
    end;
    pm := '"' + DllPath + '"';
    if IsLimited then
    begin
        Execute(fn, pm, True);
    end
    else
    begin
        Execute(fn, pm);
    end;

end;

procedure TWndSet.btnUnregClick(Sender: TObject);
var
    fn, pm:string;
    p: PWideChar;
const
	FOLDERID_SystemX86 : KNOWNFOLDERID ='{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}';
begin
    fn := 'RegSvr32';
    if IsWOW64 then
    begin
      if SHGetKnownFolderPath(FOLDERID_SystemX86, 0, 0, p) = S_OK then
      	fn := IncludeTrailingPathDelimiter(p) + 'RegSvr32';
    end;
    pm := ' /u "' + DllPath + '"';
    if IsLimited then
    begin
        Execute(fn, pm, True);
    end
    else
    begin
        Execute(fn, pm);
    end;

end;

procedure TWndSet.Load_Str(Path: string);
var
    d: string;
    ini: TMemIniFile;
begin
    Transpath := Path;
    if FileExists(Transpath) then
    begin
      ini := TMemIniFile.Create(Transpath, TEncoding.UTF8);
        try
            Font.Name := ini.ReadString('General', 'FontName', 'Arial');
            Font.Size := ini.ReadInteger('General', 'FontSize', 9);
            d := ini.ReadString('General', 'Charset', 'ASCII_CHARSET');
            d := UpperCase(d);
            if d = 'ANSI_CHARSET' then Font.Charset := 0
            else if d = 'DEFAULT_CHARSET' then Font.Charset := 1
            else if d = 'SYMBOL_CHARSET' then Font.Charset := 2
            else if d = 'MAC_CHARSET' then Font.Charset := 77
            else if d = 'SHIFTJIS_CHARSET' then Font.Charset := 128
            else if d = 'HANGEUL_CHARSET' then Font.Charset := 129
            else if d = 'JOHAB_CHARSET' then Font.Charset := 130
            else if d = 'GB2312_CHARSET' then Font.Charset := 134
            else if d = 'CHINESEBIG5_CHARSET' then Font.Charset := 136
            else if d = 'KGREEK_CHARSET' then Font.Charset := 161
            else if d = 'TURKISH_CHARSET' then Font.Charset := 162
            else if d = 'VIETNAMESE_CHARSET' then Font.Charset := 163
            else if d = 'HEBREW_CHARSET' then Font.Charset := 177
            else if d = 'ARABIC_CHARSET' then Font.Charset := 178
            else if d = 'BALTIC_CHARSET' then Font.Charset := 186
            else if d = 'RUSSIAN_CHARSET' then Font.Charset := 204
            else if d = 'THAI_CHARSET' then Font.Charset := 222
            else if d = 'EASTEUROPE_CHARSET' then Font.Charset := 238
            else if d = 'OEM_CHARSET' then Font.Charset := 255
            else Font.Charset := 0;


            Caption := ini.ReadString('SetDLG', 'DlgCaption', 'Settings');
            gbSelprop.Caption := ini.ReadString('SetDLG', 'gbSelProp', 'Use of properties');
            gbFont.Caption := ini.ReadString('SetDLG', 'gbFont', 'Font');
            gbRegDll.Caption := ini.ReadString('SetDLG', 'gbRegDll', 'IAccessible2Proxy.dll');
            btnFont.Caption := ini.ReadString('SetDLG', 'btnFont', '&Font...');
            btnReg.Caption := ini.ReadString('SetDLG', 'btnReg', '&Register');
            btnUnreg.Caption := ini.ReadString('SetDLG', 'btnUnreg', '&Unregister');
            btnOK.Caption := ini.ReadString('SetDLG', 'btnOK', '&OK');
            btnCancel.Caption := ini.ReadString('SetDLG', 'btnCancel', '&Cancel');
            cbExTip.Caption := ini.ReadString('General', 'ExTooltip', 'Expansion tooltip');
            SizeChange;

        finally
            ini.Free;
        end;
    end;
end;



procedure TWndSet.FormCreate(Sender: TObject);
var
    ini: TMemIniFile;
begin
    APPDir :=  IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
    DllPath := APPDir + 'IAccessible2Proxy.dll';
    SPath := IncludeTrailingPathDelimiter(GetMyDocPath) + 'MSAAV.ini';

    ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
    try

        Font.Name := Ini.ReadString('Settings', 'FontName', Font.Name);
            Font.Size := ini.ReadInteger('Settings', 'FontSize', Font.Size);
            Font.Charset := ini.ReadInteger('Settings', 'Charset', 0);


    finally
        ini.Free;
    end;
    M1 := Font.Name + ',' + IntToStr(DefFont);
    Memo1.Lines[0] := M1;

end;

procedure TWndSet.WMDPIChanged(var Message: TMessage);
begin
  if (DefX > 0) and (DefY > 0) then
  begin
    scaleX := Message.WParamLo / DefX;
    scaleY := Message.WParamHi / DefY;
    SizeChange;
  end;
end;

end.
