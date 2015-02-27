unit frmMSAAV;
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
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, System.UITypes,
  Dialogs, ImgList, ComCtrls, ToolWin, StdCtrls, oleacc, system.Types, ShellAPI,
  IniFiles, ActnList, CommCtrl, FocusRectWnd, ClipBrd,
  ActiveX, MSHTML_tlb, ExtCtrls, Shlobj, ComObj, MSXML2_tlb,
  MsHTMDid, SHDOCVW, global_var, get_iweb, TransCheckBox,iAccessible2Lib_tlb, ISimpleDOM,
  treelist, Menus, MMSystem, Thread, TipWnd, UIAutomationClient_TLB, UIA_TLB, PanelBtn,
  AccCTRLs, frmSet, Math, IntList, System.Actions, StrUtils;

const
    IID_IEnumVARIANT:    TGUID = '{00020404-0000-0000-c000-000000000046}';
    IID_IDispatchEx : TGuid = '{a6ef9860-c720-11d0-9337-00a0c90dcaa9}';
    IID_IServiceProvider: TGUID = '{6D5140C1-7436-11CE-8034-00AA006009FA}';
    IID_NULL : TGUID = '{00000000-0000-0000-0000-000000000000}';
    PROPID_ACC_NAME             : TGUID = '{608d3df8-8128-4aa7-a428-f55e49267291}';
    PROPID_ACC_VALUE            : TGUID = '{123fe443-211a-4615-9527-c45a7e93717a}';
    PROPID_ACC_DESCRIPTION      : TGUID = '{4d48dfe4-bd3f-491f-a648-492d6f20c588}';
    PROPID_ACC_ROLE             : TGUID = '{cb905ff2-7bd1-4c05-b3c8-e6c241364d70}';
    PROPID_ACC_STATE            : TGUID = '{a8d4d5b0-0a21-42d0-a5c0-514e984f457b}';
    PROPID_ACC_HELP             : TGUID = '{c831e11f-44db-4a99-9768-cb8f978b7231}';
    PROPID_ACC_KEYBOARDSHORTCUT : TGUID = '{7d9bceee-7d1e-4979-9382-5180f4172c34}';
    PROPID_ACC_DEFAULTACTION    : TGUID = '{180c072b-c27f-43c7-9922-f63562a4632b}';

    PROPID_ACC_HELPTOPIC        : TGUID = '{787d1379-8ede-440b-8aec-11f7bf9030b3}';
    PROPID_ACC_FOCUS            : TGUID = '{6eb335df-1c29-4127-b12c-dee9fd157f2b}';
    PROPID_ACC_SELECTION        : TGUID = '{b99d073c-d731-405b-9061-d95e8f842984}';
    PROPID_ACC_PARENT           : TGUID = '{474c22b6-ffc2-467a-b1b5-e958b4657330}';

    PROPID_ACC_NAV_UP           : TGUID = '{016e1a2b-1a4e-4767-8612-3386f66935ec}';
    PROPID_ACC_NAV_DOWN         : TGUID = '{031670ed-3cdf-48d2-9613-138f2dd8a668}';
    PROPID_ACC_NAV_LEFT         : TGUID = '{228086cb-82f1-4a39-8705-dcdc0fff92f5}';
    PROPID_ACC_NAV_RIGHT        : TGUID = '{cd211d9f-e1cb-4fe5-a77c-920b884d095b}';
    PROPID_ACC_NAV_PREV         : TGUID = '{776d3891-c73b-4480-b3f6-076a16a15af6}';
    PROPID_ACC_NAV_NEXT         : TGUID = '{1cdc5455-8cd9-4c92-a371-3939a2fe3eee}';
    PROPID_ACC_NAV_FIRSTCHILD   : TGUID = '{cfd02558-557b-4c67-84f9-2a09fce40749}';
    PROPID_ACC_NAV_LASTCHILD    : TGUID = '{302ecaa5-48d5-4f8d-b671-1a8d20a77832}';
    P2WDEF = 350;
    P4HDEF = 250;

    EM_SHOWBALLOONTIP = $1503;
    EM_HIDEBALLOONTIP = $1504;
    TTI_ERROR = 3;
    HTML_STYLE = '<head>' + #13#10 +
		'<meta charset="utf-8">' + #13#10 +
    '<style>' + #13#10 + 'body {font-family:verdana}' + #13#10 +
		'input[type=checkbox]:checked ~ section[id^="x-details"] {display:block}' + #13#10 +
		'input[type=checkbox] ~ section[id^="x-details"] {display:none}' + #13#10 +
		'input[type=checkbox]:checked ~ label:before { transform: rotate(90deg); }' + #13#10 +
		'input[type=checkbox] ~ label:before { content:"►"; font-size: 1em; position: relative;  transition: .5s linear; }' + #13#10 +
		'input[type=checkbox]:checked ~ label:before { content:"▼"; font-size: 1em; position: relative; transition: .5s linear;  color: blue; }' + #13#10 +
		'input[type=checkbox]:focus + label {outline:2px solid #4040FF;}' + #13#10 +
		'input[type=checkbox] + label {outline:1px solid grey;}' + #13#10 +
		'label {display:inline-block;width:30em;height:1em;margin-left:-24px;background:white;font-family:verdana;font-weight:bold;cursor:pointer;padding-top:.5em;}' + #13#10 +
		'section[id^="x-details"] {width:29em;border:1px dashed}' + #13#10 +
		'</style>' + #13#10 + '</head>' + #13#10 + '<body>';
type
  PEditBalloonTip = ^TEditBalloonTip;



  EDITBALLOONTIP = record

    cbStruct: DWORD;

    pszTitle: PWChar;

    pszText: PWChar;

    ttiIcon: Integer;

  end;



  TEditBalloonTip = EDITBALLOONTIP;
   PTreeData = ^TTreeData;
   TTreeData = record
      Acc: IAccessible;
      iID: integer;
   end;
   PIE = ^FIE;
    FIE = record
        Iweb: IWebBrowser2;
    end;
  //FCPList用
    PCP = ^TCP;
    TCP = record
        CP: IConnectionPoint;
        Cookie: integer;
    end;
    PUIA = ^TUIA;
    TUIA = record
        ID: integer;
        PName: string;
        Value: OleVariant;
    end;
  TwndMSAAV = class(TForm)
    ImageList1: TImageList;
    ActionList1: TActionList;
    acFocus: TAction;
    acCursor: TAction;
    acRect: TAction;
    acCopy: TAction;
    acOnlyFocus: TAction;
    Timer1: TTimer;
    acParent: TAction;
    acChild: TAction;
    acPrevS: TAction;
    acNextS: TAction;
    acHelp: TAction;
    FontDialog1: TFontDialog;
    Panel1: TPanel;
    Toolbar1: TAccToolbar;
    tbFocus: TToolButton;
    tbCursor: TToolButton;
    ToolButton3: TToolButton;
    tbRectAngle: TToolButton;
    ToolButton5: TToolButton;
    tbCopy: TToolButton;
    ToolButton2: TToolButton;
    tbOnlyFocus: TToolButton;
    ToolButton1: TToolButton;
    tbParent: TToolButton;
    tbChild: TToolButton;
    tbPrevS: TToolButton;
    tbNextS: TToolButton;
    ToolButton4: TToolButton;
    tbHelp: TToolButton;
    acMSAAMode: TAction;
    tbMSAAMode: TToolButton;
    PopupMenu1: TPopupMenu;
    mnuReg: TMenuItem;
    mnuUnreg: TMenuItem;
    ImageList2: TImageList;
    tbRegister: TToolButton;
    Splitter1: TSplitter;
    Timer2: TTimer;
    tbShowTip: TToolButton;
    acShowTip: TAction;
    MainMenu1: TMainMenu;
    mnuView: TMenuItem;
    mnuSelD: TMenuItem;
    mnuMSAA: TMenuItem;
    mnuARIA: TMenuItem;
    mnuHTML: TMenuItem;
    mnuIA2: TMenuItem;
    mnuUIA: TMenuItem;
    Panel2: TPanel;
    PB1: TPanelButton;
    Panel4: TPanel;
    PB3: TPanelButton;
    Panel5: TPanel;
    Splitter2: TSplitter;
    Panel3: TPanel;
    PB2: TPanelButton;
    acTVcol: TAction;
    acTLCol: TAction;
    acMMCol: TAction;
    mnuColl: TMenuItem;
    mnuTV: TMenuItem;
    mnuTL: TMenuItem;
    mnuCode: TMenuItem;
    Memo1: TAccMemo;
    TreeList1: TAccTreeList;
    TreeView1: TAccTreeView;
    mnuTVCont: TMenuItem;
    mnuTarget: TMenuItem;
    mnuAll: TMenuItem;
    acSetting: TAction;
    mnuBln: TMenuItem;
    mnublnMSAA: TMenuItem;
    mnublnIA2: TMenuItem;
    mnublnCode: TMenuItem;
    mnuLang: TMenuItem;
    ImageList3: TImageList;
    N1: TMenuItem;
    mnuTVSave: TMenuItem;
    SaveDlg: TSaveDialog;
    PopupMenu2: TPopupMenu;
    mnuSAll: TMenuItem;
    mnuSSel: TMenuItem;
    mnuTVSAll: TMenuItem;
    mnuTVSSel: TMenuItem;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure acFocusExecute(Sender: TObject);
    procedure acCursorExecute(Sender: TObject);
    procedure acCopyExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure acRectExecute(Sender: TObject);
    procedure acOnlyFocusExecute(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure acParentExecute(Sender: TObject);
    procedure acChildExecute(Sender: TObject);
    procedure acPrevSExecute(Sender: TObject);
    procedure acNextSExecute(Sender: TObject);
    procedure acHelpExecute(Sender: TObject);
    procedure Toolbar1Enter(Sender: TObject);
    procedure Toolbar1Exit(Sender: TObject);
    procedure acMSAAModeExecute(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure TreeList1Change(Sender: TObject; Node: TTreeNode);
    procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure TreeView1Deletion(Sender: TObject; Node: TTreeNode);
    procedure FormDestroy(Sender: TObject);
    procedure acShowTipExecute(Sender: TObject);
    procedure mnuMSAAClick(Sender: TObject);
    procedure acTVcolExecute(Sender: TObject);
    procedure acTLColExecute(Sender: TObject);
    procedure acMMColExecute(Sender: TObject);
    procedure Splitter1Moved(Sender: TObject);
    procedure Splitter2Moved(Sender: TObject);
    procedure mnuTargetClick(Sender: TObject);
    procedure mnuAllClick(Sender: TObject);
    procedure acSettingExecute(Sender: TObject);
    procedure TreeView1Hint(Sender: TObject; const Node: TTreeNode;
      var Hint: string);
    procedure TreeList1Deletion(Sender: TObject; Node: TTreeNode);
    //procedure TreeView1Addition(Sender: TObject; Node: TTreeNode);
    procedure TreeList1Click(Sender: TObject);
    procedure TreeList1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TreeView1Expanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
    procedure TreeView1Addition(Sender: TObject; Node: TTreeNode);
    procedure PopupMenu2Popup(Sender: TObject);
    procedure mnuTVSAllClick(Sender: TObject);
    procedure mnuTVSSelClick(Sender: TObject);
    procedure mnuSAllClick(Sender: TObject);
    procedure mnuSSelClick(Sender: TObject);
  private
    { Private declarations }
    Cookie: Integer;
    iFocus, iRefCnt, ShowSrcLen: integer;
    CP: IConnectionPoint;
    CPC: IConnectionPointContainer;
    IE: IWebBrowser2;
    FCPList: TList;
    FIEList: TList;
    FCURList, FCACList: TList;
    hHook: THandle;
    LangList: TStringList;
    rType, rTarg: string;
    arPT: array [0..2] of TPoint;
    lMSAA: array[0..11] of string;
    lIA2: array [0..13] of string;
    lUIA: array [0..60] of string;
    oldPT: TPoint;
    WndFocus, WndLabel, WndDesc, WndTarg: TwndFocusRect;
    WndTip: TfrmTipWnd;
    hRgn1, hRgn2, hRgn3: hRgn;
    None, ConvErr, HelpURL, DllPath, APPDir, sTrue, sFalse: string;
    Created: boolean;
    oldIEHwnd: hwnd;
    iEvSrcEle, CEle: IHTMLElement;
    SDom: ISImpleDOMNode;
    UIEle: IUIAutomationElement;
    TreeTH: TreeThread;
    Treemode: boolean;
    //UIA: IUIAutomation;
    P2W, P4H: integer;
    flgMSAA, flgIA2, flgUIA, flgUIA2: integer;

    TipText, TipTextIA2: string;
    sNode: TTreeNode;
    pNode: TTreeNode;
    sRC: TRect;
    hWndTip: THandle;
    iDefIndex: integer;
    ti: TOOLINFO;
    TransPath, SPath, TransDir: string;
    UIAuto       : IUIAutomation;
    iAcc, accRoot, DocAcc: IAccessible;
    VarParent: variant;
    iMode: integer;
    ExTip, OnTree: boolean;
    procedure ExecOnlyFocus;
    function GetSource(iDoc: IHTMLDOCUMENT): string;
    procedure GetFrame(var FIEList: TList);
    procedure ConnectIEEvent;
    procedure OnMouseEnter(Sender: TObject);
    procedure OnMouseLeave(Sender: TObject);
    function MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = false): string;
    function MSAAText4Tip(pAcc: IAccessible = nil): string;
    function MSAAText4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
    function IA2Text4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
    function HTMLText: string;
    function HTMLText4FF: string;
    function ARIAText: string;
    function ARIAText4FF: string;
    function SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
    function UIAText: string;

    procedure SizeChange;
    procedure ExecMSAAMode(sender: TObject);
    procedure ExecFFArrow(sender: TObject);
    procedure Delay_ms(ms: cardinal);
    //Thread terminate
    procedure ThDone(Sender: TObject);
    procedure ReflexID(sID: string; isEle: ISimpleDOMNODE; Labelled: boolean = true);
    procedure ShowBalloonTip(Control: TWinControl; Icon: integer; Title: string; Text: string; RC: TRect; X, Y:integer; Track:boolean = false);
    procedure SetBalloonPos(X, Y:integer);
    procedure ReflexACC(ParentNode: TTreeNode; ParentAcc: IAccessible);
    procedure GetTree(ParentNode: TTreeNode; ParentAcc: IAccessible);
    procedure ReflexIA2Rel(pNode: TTreenode; ia2: IAccessible2);
    function GetSameAcc: boolean;
    function IsSameUIElement(ia1, ia2: IAccessible): boolean;
    procedure SetTreeMode(pNode: TTreeNode);
    function AccIsNull(tAcc: IAccessible): boolean;
    procedure mnuLangChildClick(Sender: TObject);
    function LoadTranslation(Sec, Ident, Def: string):string;
    function LoadTranslation_Path(Sec, Ident, Def, Path: string):string;
    function LoadTransInt(Sec, Ident:string;  Def: integer):integer;
    procedure ReflexTV(cNode: TTreeNode; var HTML: string; var iCnt: integer; ForSel: boolean = false);
    function Get_RoleText(Acc: IAccessible; Child: integer): string;
    function GetIA2Role(iRole: integer): string;
    function GetIA2State(iRole: integer): string;
    procedure ExecCmdLine;
  protected
    { Protected declarations  }

  public
    { Public declarations }

    procedure GetNaviState(AllFalse: boolean = false);
    procedure Load;
    function GetOpeMode(CName: string): integer;
    function LoadLang: string;
    procedure ShowRectWnd(bkClr: TColor);
    procedure ShowRectWnd2(bkClr: TColor; RC:TRect);
    procedure ShowLabeledWnd(bkClr: TColor; RC:TRect);
    procedure ShowDescWnd(bkClr: TColor; RC:TRect);
    procedure ShowTargWnd(bkClr: TColor; RC:TRect);
    procedure ShowTipWnd;
    function ShowMSAAText:boolean;
    procedure WMCopyData(var Msg: TWMCopyData); Message WM_COPYDATA;
    Procedure SetAbsoluteForegroundWindow(HWND: hWnd);
    function GetWindowNameLC(Wnd: HWND): string;
    //IDispatch
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    //procedure WMNotify(var Msg:TWMNotify);message WM_NOTIFY;
  end;

var
  wndMSAAV: TwndMSAAV;
  OriginalProc: TFNWndProc;
  ac: IAccessible;
  StopWait, DMode: boolean;
  bTer, bTer2: boolean;
  rNode: TTreeNode;
  TBList: TIntegerList;
  CNames: array of string;

implementation


{$R *.dfm}



function GetLastErrorStr(ErrorCode: Integer): String;
const
  MAX_MES = 512;
var
  Buf: PChar;
begin
  Buf := AllocMem(MAX_MES);
  try
    FormatMessage(Format_Message_From_System, Nil, ErrorCode,
                  (SubLang_Default shl 10) + Lang_Neutral,
                  Buf, MAX_MES, Nil);
  finally
    Result := Buf;
    FreeMem(Buf);
  end;
end;

function IntToBoolStr(i: integer): string;
begin
    Result := 'True';
    if i = 0 then
        Result := 'False';
end;

function TruncPow(i, t:integer): integer;
begin
    Result := Trunc(IntPower(i, t));
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

function DeleteCRLF(d: widestring): widestring;
var
    s: widestring;
begin
    s := StringReplace(d, #10,' ' ,[rfReplaceAll , rfIgnoreCase ]);
    s := StringReplace(s, #13,'' ,[rfReplaceAll , rfIgnoreCase ]);
    result := s;
end;

function VarHaveValue(v: variant): boolean;
begin
    result := true;
    if VarType(v) = varEmpty then result := false;
    if VarIsEmpty(v) then result := false;
    if VarIsClear(v) then result := false;

end;

procedure ShowErr(Msg: string);
begin
    if Dmode then
        MessageDlg(Msg , mtError, [mbOK], 0);
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
    raise Exception.Create('仮想フォルダのため取得できません');
  end;
  Result := StrPas(Buffer);
end;

procedure WinEventProc(hWinEventHook: THandle; event: DWORD; hwnd: HWND;
idObject, idChild: Longint; idEventThread, dwmsEventTime: DWORD); stdcall;
var

    vChild: variant;
    i:cardinal;
    pAcc: IAccessible;
    s: string;
    pWB: IWebBrowser2;
    iDis: iDispatch;
    function SetiAcc: HResult;
    begin
        result := AccessibleObjectFromEvent(hwnd, idObject, idChild, @pAcc, vChild);
        wndMSAAV.UIAuto.GetFocusedElement(wndMSAAV.UIEle);
    end;
begin
    case event of


        EVENT_OBJECT_FOCUS:
        begin
            if not wndMSAAV.acFocus.Checked then
                Exit;
            try

            s := wndMSAAV.GetWindowNameLC(hwnd);
            if (s = 'internet explorer_server') and (wndMSAAV.oldIEHwnd <> hwnd) then
            begin
                GetIEFromHWND(hwnd, pWB);
                if pWB <> nil then
                begin
                    wndMSAAV.oldIEHwnd := hwnd;
                    wndMSAAV.IE := pWB;
                    wndMSAAV.ConnectIEEvent;
                    wndMSAAV.CEle := nil;
                    wndMSAAV.tbParent.Enabled := false;
                    wndMSAAV.tbChild.Enabled := false;
                    wndMSAAV.tbPrevS.Enabled := false;
                    wndMSAAV.tbNextS.Enabled := false;
                    wndMSAAV.acParent.Enabled := wndMSAAV.tbParent.Enabled;
                    wndMSAAV.acChild.Enabled := wndMSAAV.tbChild.Enabled;
                    wndMSAAV.acPrevS.Enabled := wndMSAAV.tbPrevS.Enabled;
                    wndMSAAV.acNextS.Enabled := wndMSAAV.tbNextS.Enabled;
                end;
            end;

            i := GetWindowLong(hwnd, GWL_HINSTANCE);
            if i = hInstance then
            begin

                exit;
            end;
            //if (not wndMSAAV.acOnlyFOcus.Checked) and (not wndMSAAV.acFocus.Checked) then exit;

            //if SUCCEEDED(AccessibleObjectFromEvent(hwnd, idObject, idChild, @wndMSAAV.iAcc, vChild)) then
            //begin

                {if (s = 'internet explorer_server') or (s = 'macromediaflashplayeractivex') then
                    wndMSAAV.iMode := 0
                else if (s = 'mozillawindowclass') or (s = 'chrome_renderwidgethosthwnd') or (s = 'mozillawindowclass') or (s = 'chrome_widgetwin_0') or (s = 'chrome_widgetwin_1') then
                    wndMSAAV.iMode := 1
                else
                    wndMSAAV.iMode := 2;   }
                wndMSAAV.iMode := wndMSAAV.GetOpeMode(s);
                //if wndMSAAV.iMode = 0 then
                if (not wndMSAAV.acMSAAMode.Checked) then
                begin
                    if (wndMSAAV.acOnlyFOcus.Checked) then
                    begin
                        if SUCCEEDED(SetiAcc) then
                        begin
                            wndMSAAV.varParent := vChild;
                            wndMSAAV.iAcc := pAcc;
                            //iDis := pAcc.accParent;
                            pAcc.Get_accParent(iDis);
                            wndMSAAV.accRoot := iDis as IAccessible;
                            wndMSAAV.ShowRectWnd(clYellow);
                        end;

                    end
                    else
                    begin
                        if wndMSAAV.iMode <> 2 then
                        begin
                            if (wndMSAAV.acFocus.Checked) then
                            begin
                                if SUCCEEDED(SetiAcc) then
                                begin

                                    wndMSAAV.varParent := vChild;
                                    wndMSAAV.iAcc := pAcc;
                                    pAcc.Get_accParent(iDis);
                                    wndMSAAV.accRoot := iDis as IAccessible;
                                    if (wndMSAAV.acRect.Checked) then
                                    begin
                                        wndMSAAV.ShowRectWnd(clRed);
                                    end;

                                    wndMSAAV.Timer2.Enabled := false;
                                    wndMSAAV.Timer2.Enabled := True;
                                end;
                            end;
                        end;
                    end;
                end
                else
                begin
                    if (wndMSAAV.acFocus.Checked) then
                    begin
                        if SUCCEEDED(SetiAcc) then
                        begin
                            wndMSAAV.varParent := vChild;
                            wndMSAAV.iAcc := pAcc;
                            pAcc.Get_accParent(iDis);
                            wndMSAAV.accRoot := iDis as IAccessible;
                            if (wndMSAAV.acRect.Checked) then
                            begin
                                wndMSAAV.ShowRectWnd(clRed);
                            end;
                            wndMSAAV.Timer2.Enabled := false;
                            wndMSAAV.Timer2.Enabled := True;
                        end;
                    end;
                end;

            except
                on E:Exception do
                    ShowErr(E.Message);
            end;
        end;

    end;
end;


function TwndMSAAV.Get_RoleText(Acc: IAccessible; Child: integer): string;
var
    PC:PChar;
    ovValue, ovChild: OleVariant;
begin
    ovChild := Child;
    //ovValue := Acc.accRole[ovChild];
    Acc.Get_accRole(ovChild, ovValue);
    Result := None;
    try
                    //if (IDispatch(ovValue) = nil) then

        if VarHaveValue(ovValue) then
        begin
            //GetMem(PC,255);
            if VarIsNumeric(ovValue) then
            begin
                PC := StrAlloc(255);
                GetRoleTextW(ovValue, PC, StrBufSize(PC));
                Result := PC;
                StrDispose(PC);
            end
            else if VarIsStr(ovValue) then
            begin
                Result := VarToStr(ovValue);
            end;
        end;
    except
        on E:Exception do
            ShowErr(E.Message);

    end;
end;

function TwndMSAAV.GetIA2Role(iRole: integer): string;
var
	ini: TMemIniFile;
  PC:PChar;
  ovValue: OleVariant;

begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
        try
            case iRole of
            IA2_ROLE_CANVAS: result := ini.ReadString('IA2', 'IA2_ROLE_CANVAS', 'canvas');
            IA2_ROLE_CAPTION: result := ini.ReadString('IA2', 'IA2_ROLE_CAPTION', 'Caption');
            IA2_ROLE_CHECK_MENU_ITEM: result := ini.ReadString('IA2', 'IA2_ROLE_CHECK_MENU_ITEM', 'Check menu item');
            IA2_ROLE_COLOR_CHOOSER: result := ini.ReadString('IA2', 'IA2_ROLE_COLOR_CHOOSER', 'Color chooser');
            IA2_ROLE_DATE_EDITOR: result := ini.ReadString('IA2', 'IA2_ROLE_DATE_EDITOR', 'Date editor');
            IA2_ROLE_DESKTOP_ICON: result := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_ICON', 'Desktop icon');
            IA2_ROLE_DESKTOP_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_PANE', 'Desktop pane');
            IA2_ROLE_DIRECTORY_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_DIRECTORY_PANE', 'Directory pane');
            IA2_ROLE_EDITBAR: result := ini.ReadString('IA2', 'IA2_ROLE_EDITBAR', 'Editbar');
            IA2_ROLE_EMBEDDED_OBJECT: result := ini.ReadString('IA2', 'IA2_ROLE_EMBEDDED_OBJECT', 'Embedded object');
            IA2_ROLE_ENDNOTE: result := ini.ReadString('IA2', 'IA2_ROLE_ENDNOTE', 'Endnote');
            IA2_ROLE_FILE_CHOOSER: result := ini.ReadString('IA2', 'IA2_ROLE_FILE_CHOOSER', 'File chooser');
            IA2_ROLE_FONT_CHOOSER: result := ini.ReadString('IA2', 'IA2_ROLE_FONT_CHOOSER', 'Font chooser');
            IA2_ROLE_FOOTER: result := ini.ReadString('IA2', 'IA2_ROLE_FOOTER', 'Footer');
            IA2_ROLE_FOOTNOTE: result := ini.ReadString('IA2', 'IA2_ROLE_FOOTNOTE', 'Footnote');
            IA2_ROLE_FORM: result := ini.ReadString('IA2', 'IA2_ROLE_FORM', 'Form');
            IA2_ROLE_FRAME: result := ini.ReadString('IA2', 'IA2_ROLE_FRAME', 'Frame');
            IA2_ROLE_GLASS_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_GLASS_PANE', 'Glass pane');
            IA2_ROLE_HEADER: result := ini.ReadString('IA2', 'IA2_ROLE_HEADER', 'Header');
            IA2_ROLE_HEADING: result := ini.ReadString('IA2', 'IA2_ROLE_HEADING', 'Heading');
            IA2_ROLE_ICON: result := ini.ReadString('IA2', 'IA2_ROLE_ICON', 'Icon');
            IA2_ROLE_IMAGE_MAP: result := ini.ReadString('IA2', 'IA2_ROLE_IMAGE_MAP', 'Image map');
            IA2_ROLE_INPUT_METHOD_WINDOW: result := ini.ReadString('IA2', 'IA2_ROLE_INPUT_METHOD_WINDOW', 'Input method window');
            IA2_ROLE_INTERNAL_FRAME: result := ini.ReadString('IA2', 'IA2_ROLE_INTERNAL_FRAME', 'Internal frame');
            IA2_ROLE_LABEL: result := ini.ReadString('IA2', 'IA2_ROLE_LABEL', 'Label');
            IA2_ROLE_LAYERED_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_LAYERED_PANE', 'Layered pane');
            IA2_ROLE_NOTE: result := ini.ReadString('IA2', 'IA2_ROLE_NOTE', 'Note');
            IA2_ROLE_OPTION_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_OPTION_PANE', 'Option pane');
            IA2_ROLE_PAGE: result := ini.ReadString('IA2', 'IA2_ROLE_PAGE', 'Role pane');
            IA2_ROLE_PARAGRAPH: result := ini.ReadString('IA2', 'IA2_ROLE_PARAGRAPH', 'Paragraph');
            IA2_ROLE_RADIO_MENU_ITEM: result := ini.ReadString('IA2', 'IA2_ROLE_RADIO_MENU_ITEM', 'Radio menu item');
            IA2_ROLE_REDUNDANT_OBJECT: result := ini.ReadString('IA2', 'IA2_ROLE_REDUNDANT_OBJECT', 'Redundant object');
            IA2_ROLE_ROOT_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_ROOT_PANE', 'Root pane');
            IA2_ROLE_RULER: result := ini.ReadString('IA2', 'IA2_ROLE_RULER', 'Ruler');
            IA2_ROLE_SCROLL_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_SCROLL_PANE', 'Scroll pane');
            IA2_ROLE_SECTION: result := ini.ReadString('IA2', 'IA2_ROLE_SECTION', 'Section');
            IA2_ROLE_SHAPE: result := ini.ReadString('IA2', 'IA2_ROLE_SHAPE', 'Shape');
            IA2_ROLE_SPLIT_PANE: result := ini.ReadString('IA2', 'IA2_ROLE_SPLIT_PANE', 'Split pane');
            IA2_ROLE_TEAR_OFF_MENU: result := ini.ReadString('IA2', 'IA2_ROLE_TEAR_OFF_MENU', 'Tear off menu');
            IA2_ROLE_TERMINAL: result := ini.ReadString('IA2', 'IA2_ROLE_TERMINAL', 'Terminal');
            IA2_ROLE_TEXT_FRAME: result := ini.ReadString('IA2', 'IA2_ROLE_TEXT_FRAME', 'Text frame');
            IA2_ROLE_TOGGLE_BUTTON: result := ini.ReadString('IA2', 'IA2_ROLE_TOGGLE_BUTTON', 'Toggle button');
            IA2_ROLE_VIEW_PORT: result := ini.ReadString('IA2', 'IA2_ROLE_VIEW_PORT', 'View port');
            else
            begin
                PC := StrAlloc(255);
                ovValue := iROle;
                GetRoleTextW(ovValue, PC, StrBufSize(PC));
                result := PC;
                StrDispose(PC);
            end;
            end;
        finally
            ini.Free;
        end;
end;

function TwndMSAAV.GetIA2State(iRole: integer): string;
var
	ini: TMemIniFile;
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
        try
            if (iRole and IA2_STATE_ACTIVE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_ACTIVE', 'Active');
            end;

            if (iRole and IA2_STATE_ARMED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_ARMED', 'Armed');
            end;

            if (iRole and IA2_STATE_DEFUNCT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_DEFUNCT', 'Defunct');
            end;

            if (iRole and IA2_STATE_EDITABLE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_EDITABLE', 'Editable');
            end;

            if (iRole and IA2_STATE_HORIZONTAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_HORIZONTAL', 'Horizontal');
            end;

            if (iRole and IA2_STATE_ICONIFIED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_ICONIFIED', 'Iconified');
            end;

            if (iRole and IA2_STATE_INVALID_ENTRY) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_INVALID_ENTRY', 'Invalid entry');
            end;

            if (iRole and IA2_STATE_MANAGES_DESCENDANTS) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_MANAGES_DESCENDANTS', 'Manages descendants');
            end;

            if (iRole and IA2_STATE_MODAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_MODAL', 'Modal');
            end;

            if (iRole and IA2_STATE_MULTI_LINE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_MULTI_LINE', 'Multi line');
            end;

            if (iRole and IA2_STATE_OPAQUE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_OPAQUE', 'Opaque');
            end;

            if (iRole and IA2_STATE_REQUIRED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_REQUIRED', 'Required');
            end;

            if (iRole and IA2_STATE_SELECTABLE_TEXT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_SELECTABLE_TEXT', 'Selectable text');
            end;

            if (iRole and IA2_STATE_SINGLE_LINE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_SINGLE_LINE', 'Single line');
            end;

            if (iRole and IA2_STATE_STALE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_STALE', 'Stale');
            end;

            if (iRole and IA2_STATE_SUPPORTS_AUTOCOMPLETION) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_SUPPORTS_AUTOCOMPLETION', 'Supports autocompletion');
            end;

            if (iRole and IA2_STATE_TRANSIENT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_TRANSIENT', 'Transient');
            end;

            if (iRole and IA2_STATE_VERTICAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + ini.ReadString('IA2', 'IA2_STATE_VERTICAL', 'Vertical');
            end;

        finally
            ini.Free;
        end;
end;

function TwndMSAAV.LoadTransInt(Sec, Ident:string; Def: integer):integer;
var
	ini: TMemIniFile;
    s: string;
begin
    Result := Def;
    ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	try
        if ini.SectionExists(Sec) then
            Result := Ini.Readinteger(Sec, Ident, Def)
        else
        begin
            if LowerCase(RightStr(Sec, 1)) = 's' then
            begin
                s := LeftStr(Sec, Length(Sec)-1);
                if ini.SectionExists(s) then
                    Result := Ini.Readinteger(S, Ident, Def);
            end
            else
            begin
                s := Sec + 's';
                if ini.SectionExists(s) then
                    Result := Ini.Readinteger(S, Ident, Def);
            end;
        end;
    finally
        ini.Free;
    end;
end;

function TwndMSAAV.LoadTranslation_Path(Sec, Ident, Def, Path: string):string;
var
	ini: TMemIniFile;
    s: string;
begin
    Result := Def;
    ini := TMemIniFile.Create(Path, TEncoding.Unicode);
	try
        if ini.SectionExists(Sec) then
            Result := Ini.ReadString(Sec, Ident, Def)
        else
        begin
            if LowerCase(RightStr(Sec, 1)) = 's' then
            begin
                s := LeftStr(Sec, Length(Sec)-1);
                if ini.SectionExists(s) then
                    Result := Ini.ReadString(S, Ident, Def);
            end
            else
            begin
                s := Sec + 's';
                if ini.SectionExists(s) then
                    Result := Ini.ReadString(S, Ident, Def);
            end;
        end;
    finally
        ini.Free;
    end;

end;

function TwndMSAAV.LoadTranslation(Sec, Ident, Def: string):string;
var
	ini: TMemIniFile;
    s: string;
begin
    Result := Def;
    ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	try
        if ini.SectionExists(Sec) then
            Result := Ini.ReadString(Sec, Ident, Def)
        else
        begin
            if LowerCase(RightStr(Sec, 1)) = 's' then
            begin
                s := LeftStr(Sec, Length(Sec)-1);
                if ini.SectionExists(s) then
                    Result := Ini.ReadString(S, Ident, Def);
            end
            else
            begin
                s := Sec + 's';
                if ini.SectionExists(s) then
                    Result := Ini.ReadString(S, Ident, Def);
            end;
        end;
    finally
        ini.Free;
    end;
end;

procedure TwndMSAAV.Delay_ms(ms: cardinal);
var
    STime: Cardinal;
begin


    STime := timeGetTime;
    while True do
    begin
        if StopWait then break;
        if (timeGetTime - STime) >= ms then break;

        Application.ProcessMessages;
        Sleep(1);
    end;
    StopWait := false;
end;

procedure ReflexList(var FIEList: TList; hDoc: IHTMLDocument2);
var
    iLen, i: Integer;
    iDoc: IHTMLDocument2;
    iContainer: IOLEContainer;
    enumerator: ActiveX.IEnumUnknown;
    unkFrame: IUnknown;
    nFetched: PLongInt;
    iWeb: IWebBrowser2;
    TIE: PIE;
const
    IID_IOleContainer: TGUID = '{0000011b-0000-0000-C000-000000000046}';
begin
    iLen := hDoc.frames.length;
    if iLen = 0 then Exit;
    if SUCCEEDED(hDoc.QueryInterface(IID_IOleContainer, iContainer)) then
    begin
        if SUCCEEDED(iContainer.EnumObjects(OLECONTF_EMBEDDINGS or OLECONTF_OTHERS, enumerator)) then
        begin
            for i := 0 to iLen - 1 do
            begin
                nFetched := nil;
                if SUCCEEDED(enumerator.Next(1, unkFrame, nFetched)) then
                begin
                    if SUCCEEDED(unkframe.QueryInterface(IID_IWebBrowser2, iWeb)) then
                    begin
                        New(TIE);
                        TIE^.Iweb := iWeb;
                        FIEList.Add(TIE);
                        if SUCCEEDED(iWeb.Document.QueryInterface(IID_IHTMLDocument2, iDoc)) then
                        begin
                            iLen := iDoc.frames.length;
                            if iLen > 0 then
                                ReflexList(FIEList, iDoc);
                        end;
                    end;
                end;
            end;
        end;
    end;
end;


procedure TwndMSAAV.GetFrame(var FIEList: TList);
var
    TIE: PIE;
    iDoc: IHTMLDocument2;
    i: integer;
begin
    if IE = nil then Exit;
    for i := 0 to FIEList.Count - 1 do
    begin
        if Assigned(FIEList.Items[i]) then
            Dispose(FIEList.Items[i]);
    end;
    FIEList.Clear;
    New(TIE);
    TIE.Iweb := IE;
    FIEList.Add(TIE);
    if SUCCEEDED(IE.Document.QueryInterface(IID_IHTMLDocument2, iDoc)) then
    begin
        ReflexList(FIEList, iDoc);
    end;
end;

procedure TwndMSAAV.ConnectIEEvent;
var
    i: integer;
    TCP: PCP;
    iDoc: IHTMLDocument2;
    pWB: IWebBrowser2;
    iDis: IDispatch;
begin
    if CP <> nil then
        CP.Unadvise(Cookie);
    if SUCCEEDED(IE.QueryInterface(IID_IConnectionPointContainer, CPC)) then
    begin
        if SUCCEEDED(CPC.FindConnectionPoint(DWebBrowserEvents2, CP)) then
        begin
            CP.Advise(Self, Cookie);
        end;

    end;
     for i := 0 to FCPList.Count - 1 do
     begin
        if Assigned(FCPList.Items[i]) then
        begin
            PCP(FCPList.Items[i])^.CP.Unadvise(PCP(FCPList.Items[i])^.Cookie);
            Dispose(FCPList.Items[i]);
        end;
     end;
     FCPList.Clear;
     GetFrame(FIEList);
     for i := 0 to FIEList.Count - 1 do
     begin
        pWB := PIE(FIEList.Items[i]).Iweb;
        if (pWB.LocationURL <> '') then
        begin
            if SUCCEEDED(pWB.Document.QueryInterface(IID_IHTMLDocument, iDis)) then
            begin
                if SUCCEEDED(iDis.QueryInterface(IID_IHTMLDocument2, iDoc)) then
                begin
                    if SUCCEEDED(iDis.QueryInterface(IID_IConnectionPointContainer, CPC)) then
                    begin
                        New(TCP);
                        if SUCCEEDED(CPC.FindConnectionPoint(DIID_HTMLDocumentEvents, TCP^.CP)) then
                        begin

                            TCP^.CP.Advise(Self, TCP^.Cookie);
                            FCPList.Add(TCP);
                        end
                        else
                            Dispose(TCP);

                    end;
                end;
            end;
        end;
     end;
end;

Procedure TwndMSAAV.SetAbsoluteForegroundWindow(HWND: hWnd);
var
    nTargetID, nForegroundID : Integer;
    sp_time: Integer;
begin
    nForegroundID := GetWindowThreadProcessId(GetForegroundWindow, nil);
    nTargetID := GetWindowThreadProcessId(hWnd, nil );
    AttachThreadInput(nTargetID, nForegroundID, TRUE );
    SystemParametersInfo( SPI_GETFOREGROUNDLOCKTIMEOUT,0,@sp_time,0);
    SystemParametersInfo( SPI_SETFOREGROUNDLOCKTIMEOUT,0,Pointer(0),0);
    Application.ProcessMessages;
    SetForegroundWindow(hWnd);
    SystemParametersInfo( SPI_SETFOREGROUNDLOCKTIMEOUT,0,@sp_time,0);
    AttachThreadInput(nTargetID, nForegroundID, FALSE );

end;

procedure TwndMSAAV.ExecCmdLine;
var
    i: Integer;
    d: string;
begin
    if ParamCount > 0 then
    begin
        DMode := False;
        for i := 1 to ParamCount do
        begin
          d := LowerCase(ParamStr(i));
          if (d = '-fronly') or (d = '/fronly') then
            begin

                acOnlyFocus.Checked := True;
                tbOnlyFocus.Down := True;
                ExecOnlyFocus;
            end
            else if (d = '-d') or (d = '/d') then
              DMode := True;
        end;


    end;
end;

procedure TwndMSAAV.WMCopyData(var Msg: TWMCopyData);
var
    List: TStringList;
    i: integer;
    d: string;
begin
    if msg.CopyDataStruct.dwData=$00000325 then
    begin
      List := TStringList.Create;
      try
          SetString(d, PChar(Msg.CopyDataStruct.lpData),
            Msg.CopyDataStruct.cbData div SizeOf(Char));
          List.CommaText := d;
          for i := 0 to List.Count - 1 do
          begin
            d := LowerCase(List[i]);
            if (d = '-fronly') or (d = '/fronly') then
            begin

                acOnlyFocus.Checked := True;
                tbOnlyFocus.Down := True;
                ExecOnlyFocus;
            end
            else if (d = '-d') or (d = '/d') then
              DMode := True;
        end;
      finally
        List.Free;
      end;
    end;
end;

function GetMultiState(State: cardinal): widestring;
var
    PC:PChar;
    List: TStringList;
    i: integer;
begin
    List := TStringList.Create;
    try
    if (State and STATE_SYSTEM_ALERT_HIGH) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_ALERT_HIGH, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ALERT_MEDIUM) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_ALERT_MEDIUM, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ALERT_LOW) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_ALERT_LOW, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ANIMATED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_ANIMATED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_BUSY) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_BUSY, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_CHECKED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_CHECKED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_COLLAPSED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_COLLAPSED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_DEFAULT) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_DEFAULT, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_EXPANDED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_EXPANDED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_EXTSELECTABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_EXTSELECTABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FLOATING) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_FLOATING, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FOCUSABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_FOCUSABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FOCUSED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_FOCUSED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_HASPOPUP) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_HASPOPUP, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_HOTTRACKED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_HOTTRACKED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_INVISIBLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_INVISIBLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_LINKED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_LINKED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MARQUEED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_MARQUEED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MIXED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_MIXED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MOVEABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_MOVEABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MULTISELECTABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_MULTISELECTABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    //if (State and STATE_SYSTEM_NORMAL) <> 0 then
    if State  = 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_NORMAL, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_OFFSCREEN) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_OFFSCREEN, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_PRESSED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_PRESSED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_PROTECTED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_PROTECTED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_READONLY) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_READONLY, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELECTABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_SELECTABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELECTED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_SELECTED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELFVOICING) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_SELFVOICING, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SIZEABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_SIZEABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_TRAVERSED) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_TRAVERSED, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_UNAVAILABLE) <> 0 then
    begin
        PC := StrAlloc(255);
        GetStateTextW(STATE_SYSTEM_UNAVAILABLE, PC, StrBufSize(PC));
        List.Add(PC);
        StrDispose(PC);
    end;
        for I := 0 to List.Count - 1 do
        begin
            if i = 0 then
                Result := List.Strings[i]
            else
                Result := Result + ' , ' + List.Strings[i];
        end;
    finally
        List.Free;
    end;

end;

procedure TwndMSAAV.GetTree(ParentNode: TTreeNode; ParentAcc: IAccessible);
begin
    //TreeView1.Items.BeginUpdate;
    try
    OnTree := True;
    bTer := False;
    sNode := nil;
    ReflexACC(ParentNode, ParentAcc);

    if bTer then
    begin
        sNode := nil;
        TreeView1.Items.Clear;
    end;


        OnTree := False;
        if Assigned(sNode) then
        begin


            if wndMSAAV.TreeView1.Items.Count > 0 then
            begin
                if sNode <> nil then
                begin
                    sNode.Expanded := true;
                    wndMSAAV.TreeView1.SetFocus;
                    wndMSAAV.TreeView1.TopItem := sNode;
                    sNode.Selected := True;
                    wndMSAAV.GetNaviState;
                end;
            end;

        end
        else
            wndMSAAV.TreeView1.Items.Clear;
    finally
        //TreeView1.Items.EndUpdate;
    end;
end;

procedure TwndMSAAV.ReflexACC(ParentNode: TTreeNode; ParentAcc: IAccessible);
var
    cAcc, tAcc: IAccessible;
    iChild, i, iCH: integer;
    cNode: TTreeNode;
    ws: Widestring;
    oc: OleVariant;
    iDis: iDispatch;
    Role, NodeCap: string;
    TD: PTreeData;
    cRC: TRect;

    ovChild: Olevariant;
    bOK: boolean;

     procedure AddChild;
    begin
        New(TD);
        TD^.Acc := cAcc;
        TD^.iID := iCH;
        pNode := cNode;
        rNode := TreeView1.Items.AddChildObject(pNode, NodeCap, Pointer(TD));
    end;

    procedure SName;
    begin
        cAcc.Get_accName(ovChild, ws);
        if ws = '' then ws := None;
        Role := Get_ROLETExt(cAcc, ovChild);
    end;

    procedure SCnt;
    begin
        cAcc.Get_accChildCount(iChild);
    end;

    procedure SLoc;
    begin
        cAcc.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, 0);
    end;

    procedure SCld;
        begin
            cAcc.Get_accChild(oc, iDis);
            tAcc := iDis as IAccessible;
            if tAcc <> nil then
                bOK := True;
        end;
begin
    try
    if bTer then Exit;
    cNode := ParentNode;
        ParentAcc.Get_accName(0, ws);
        if ws = '' then
            ws := None;
        cAcc := ParentAcc;
        Role := Get_RoleText(cAcc, 0);
        NodeCap := ws + ' - ' + Role;

        iCH := 0;
        AddChild;
        cAcc.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, 0);
        if (sRC.Left = cRC.Left) and (sRC.Top = cRC.Top) and
          (sRC.Right = cRC.Right) and (sRC.Bottom = cRC.Bottom) then
        begin

            sNode := rNode;
        end;

        cNode := rNode;

        // iChild := cAcc.accChildCount;
        cAcc.Get_accChildCount(iChild);
        for i := 0 to iChild - 1 do
        begin
            Application.ProcessMessages;
            if bTer then
                break;
            ws := '';
            iDis := nil;
            oc := i + 1;
            // iDis := cAcc.accChild[oc];
            cAcc.Get_accChild(oc, iDis);
            if iDis <> nil then
            begin
                ReflexACC(cNode, iDis as IAccessible);
            end
            else
            begin

                try
                    // ws := cAcc.accName[i + 1];
                    cAcc.Get_accName(i + 1, ws);
                except
                    ws := '';
                end;
                if ws = '' then
                    ws := None;
                Role := Get_RoleText(cAcc, i + 1);
                NodeCap := ws + ' - ' + Role;
                // pNode := cNode;
                iCH := i + 1;
                AddChild;

            end;
        end;

    except
        on E:Exception do
        begin
            ShowErr(E.Message);
        end;
    end;
end;

function TwndMSAAV.ShowMSAAText:boolean;
var
    vChild: variant;
    s: string;
    Wnd: hwnd;
    iSP: IServiceProvider;
    iEle: IHTMLElement;
    iDom: IHTMLDOMNode;
    Path: string;
    RC: TRect;
    iSEle: ISimpleDOMNode;
    SI: Smallint;
    PU, PU2: PUINT;
    WD: Word;
    PC, PC2:PChar;
    pAcc: IAccessible;
    iRes: integer;
    iDoc: IHTMLDocument2;
    pEle, TempEle: ISimpleDOMNode;
    i: integer;
    AWnd: HWND;
    bSame: boolean;
const
    Class_ISimpleDOMNode : TGUID = '{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}';
begin
    (*
    MSAA
Name:
Value:
Role:
State:
Description:

ARIA (from) DOM
Role:
Any assigned aria* attribute

HTML (from) DOM
Element name
Any assigned attributes
DOM tree (code) of selected element(s)
Example:
<p><strong>element</strong> under cursor</p>
*)
    Result := False;
    if not Assigned(iAcc) then Exit;

    {if OnTree then
    begin
        Terminated := True;
        while Ontree do
            Application.ProcessMessages;
    end;  }

    TreeList1.Items.BeginUpdate;
    TreeList1.Items.Clear;
    TreeList1.Items.EndUpdate;
    mnuTVSAll.Enabled := False;
		mnuTVSSel.Enabled := False;
    iDefIndex := -1;
    //if not Treemode then
    //begin
        //TreeView1.Items.Clear;
        //NotifyWinEvent(EVENT_OBJECT_DESTROY{EVENT_OBJECT_STATECHANGE}, TreeView1.Handle, OBJID_CLIENT, 0);
    //end;
    Memo1.Lines.Clear;
    try
        if SUCCEEDED(WindowFromAccessibleObject(iAcc, Wnd)) then
        begin

            GetNaviState(true);
            s := GetWindowNameLC(Wnd);

            CEle := nil;
            Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
            iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
            if (not acMSAAMode.Checked) then
            begin

                if ((iMode = 0) or (iMode = 1)) then
                begin
                    //if not Treemode then
                    //begin
                        if Assigned(TreeTH) then
                        begin
                            TreeTH.Terminate;
                            TreeTH.WaitFor;
                            TreeTH.Free;
                            TreeTH := nil;
                        end;


                    //end;
                end;
								if (iMode = 0) then//IE
								begin
									iSP := iAcc as IServiceProvider;
									if SUCCEEDED(iAcc.QueryInterface(IID_IServiceProvider, iSP)) then
									begin
										if (SUCCEEDED(iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle))) then
										begin
                    	if mnuMSAA.Checked then MSAAText;//MSAA text start
                          //iSP := iAcc as IServiceProvider;
													//if (SUCCEEDED(iSP.QueryService(IID_IHTMLELEMENT, IID_IHTMLELEMENT, iEle))) then
													//begin
														if (LowerCase(iEle.tagName) = 'body') and (varParent > 0) then iEle := iEVSrcEle;
                        		iDom := iEle as IHTMLDOMNODE;
                        		CEle := iEle;
                        		if mnuARIA.Checked then
                        		begin
                           	 	ARIAText;
                        		end;
                        		if mnuHTML.Checked then
                        		begin
                            		HTMLText;
                        		end;
                         	//end;
                    	if mnuUIA.Checked then UIAText;
											if (not Treemode) and (Splitter1.Enabled = True) then
											begin
												iDoc := iEle.document as IHTMLDocument2;
												if mnuAll.Checked then iEle := iDoc.body;
												if iEle <> nil then
												begin
                        	iSP := iEle as IServiceProvider;
													if (SUCCEEDED(iSP.QueryService(IID_IACCESSIBLE,IID_IACCESSIBLE, pAcc))) then
													begin
                          	//pAcc := iAcc;
														bSame := false;
														if (mnuAll.Checked) and (IsSameUIElement(DocAcc, pAcc)) then
														begin
															if TreeView1.Items.Count > 0 then
															begin
																Timer1.Enabled := false;
																try
																	bSame := GetSameAcc;
																finally
																	Timer1.Enabled := True;
																end;
															end;
														end;
														if not bSame then
														begin

                            	DocAcc := pAcc;
                              iAcc.accLocation(sRC.Left, sRC.Top,
                              sRC.Right, sRC.Bottom, 0);
                              WindowFromAccessibleObject(iAcc, AWnd);
                              TBList.Clear;
                              TreeView1.Items.Clear;
                              TreeTH := TreeThread.Create(iAcc, pAcc,
                              AWnd, iMode, None, True,
                              mnuAll.Checked);
                              TreeTH.OnTerminate := ThDone;
                              TreeTH.Start;
														end;
												 end;

                      	end;
                    	end;

                  	end;
                  end;
                end
                else if iMode = 1 then//FF
                begin
                  if mnuMSAA.Checked then // MSAA text start
                  begin
                    MSAAText;
                  end;
                  if mnuIA2.Checked then
                    SetIA2Text;
                  // iSP := iAcc as IServiceProvider;
									//if SUCCEEDED(iAcc.QueryInterface(IID_IServiceProvider, iSP)) then
                  //begin
                    iSP := iAcc as IServiceProvider;
                    GetNaviState(True);

                    iRes := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle);
										if SUCCEEDED(iRes) then
                    begin
                      SDom := isEle;
                      // GetNaviState;
                      if mnuARIA.Checked then ARIAText;
                      if mnuHTML.Checked then HTMLText4FF;
											if (not Treemode) and (Splitter1.Enabled = True) then
                      begin

												if SUCCEEDED(iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle)) then
                      	begin

                      		isEle.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);
                        	TempEle := isEle;
                        	i := 0;
													if mnuAll.Checked then
                        	begin
                              // while (WD <> NODETYPE_DOCUMENT) do
														while (LowerCase(String(PC)) <> '#document') do
                            begin
                            	Application.ProcessMessages;
                            	TempEle.get_parentNode(pEle);
                            	pEle.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);
                            	TempEle := pEle;
                            	if String(PC) = '#document' then
                            	begin
                            		isEle := pEle;
                            		break;
                            	end;
                            	inc(i);
                            	if i > 500 then
                            	begin
                            		isEle := nil;
                            		break;
                            	end;
                            // ROLE_SYSTEM_DOCUMENT
                            end; //end while
                          end;
													if isEle <> nil then
                          begin
                          	iSP := isEle as IServiceProvider;
														if (SUCCEEDED(iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, pAcc))) then
                            begin
                            	//pAcc := iAcc;
                              bSame := false;
                              if (mnuAll.Checked) and (IsSameUIElement(DocAcc, pAcc)) then
                              begin
                                Timer1.Enabled := false;
                                try
                                	bSame := GetSameAcc;
                                finally
                                	Timer1.Enabled := True;
                                end;
                              end;

                              if not bSame then
                              begin
                              	DocAcc := pAcc;
                              	iAcc.accLocation(sRC.Left, sRC.Top, sRC.Right, sRC.Bottom, 0);
                              	WindowFromAccessibleObject(iAcc, AWnd);
                              	TreeView1.Items.Clear;
                              	TreeTH := TreeThread.Create(iAcc, pAcc, AWnd, iMode, None, True, mnuAll.Checked);
                              	TreeTH.OnTerminate := ThDone;
                              	TreeTH.Start;
                              end;

                            end;
                          end;
                        //end;
                      end;
                    end
                    else // Flash
                    begin
                      SDom := nil;
                      ShowErr(GetLastErrorStr(iRes) + '(' + intToHex(iRes,
                          2) + ')');
                    end;
                    if mnuUIA.Checked then UIAText;
                  end;
                end; // FF mode end
              end
              else
              begin
                if (mnuMSAA.Checked) and (acMSAAMode.Checked) then //MSAA text start
                begin

                    MSAAText;
                    SetIA2Text;
                    GetNaviState;
                    //if cbUIA.Checked then UIAText;
                end;
                if mnuUIA.Checked then UIAText;
            end;
            if TreeList1.Items.Count > 0 then
            begin
                TreeList1.Items[0].Expand(True);
                TreeList1.Items[0].Selected := True;
                GetNaviState;
            end;
            Result := True;
        end;
    except
        on E:Exception do
            ShowErr(E.Message);
    end;
end;



procedure TwndMSAAV.ReflexIA2Rel(pNode: TTreenode; ia2: IAccessible2);
var
    isp: iserviceprovider;
    iInter: IInterface;
    ia2Targ: iaccessible2;
    iaTarg: IAccessible;
    hr, iRole, t, p, iTarg, iRel: integer;
    cNode, ccNode, cccNode: TTreeNode;
    nodes: TTreeNodes;
    iAL: IAccessibleRelation;
    s, ws: widestring;
    TD: PTreeData;

begin
    nodes := TreeList1.Items;
    Inc(iRefCnt);
    if iRefCnt > 5 then
        exit;
    try
        if SUCCEEDED(ia2.Get_nRelations(iRole)) then
        begin
            for p := 0 to iRole - 1 do
            begin
                ia2.Get_relation(p, iAL);
                cnode := nodes.AddChild(pNode, inttostr(p));

                iAL.Get_RelationType(s);
                ccNode := nodes.AddChild(cNode, rType);
                TreeList1.SetNodeColumn(ccnode, 1, s);
                ccNode := nodes.AddChild(cNode, rTarg);

                iAL.Get_nTargets(iTarg);
                for t := 0 to iTarg do
                begin
                    iInter := nil;
                    hr := iAL.Get_target(t, iInter);
                    if SUCCEEDED(hr) then
                    begin
                        hr := iInter.QueryInterface(IID_IACCESSIBLE2, ia2Targ);
                        if SUCCEEDED(hr) then
                        begin
                            hr := ia2Targ.QueryInterface
                              (IID_IServiceProvider, iSP);
                            if SUCCEEDED(hr) then
                            begin
                                hr := iSP.QueryService(IID_IACCESSIBLE,
                                    IID_IACCESSIBLE, iaTarg);
                                if SUCCEEDED(hr) then
                                begin

                                    // s := iaTarg.accName[0];
                                    iaTarg.Get_accName(0, s);
                                    ia2Targ.Role(iRole);
                                    ws := GetIA2Role(iRole);
                                    // iaTarg.Get_accValue(0, ws);
                                    cccNode :=
                                      nodes.AddChild(ccNode, inttostr(t));
                                    TreeList1.SetNodeColumn(cccNode, 1,
                                        s + ' - ' + ws);
                                    New(TD);
                                    TD^.Acc := iaTarg;
                                    TD^.iID := 0;
                                    cccNode.Data := Pointer(TD);
                                    if SUCCEEDED(ia2Targ.Get_nRelations(iRel))
                                    then
                                    begin
                                        if iRel > 0 then
                                        begin
                                        ReflexIA2Rel(ccNode, ia2Targ);
                                        end;
                                    end;
                                end;
                            end;
                        end;
                    end;
                end;
                { iAL.Get_RelationType(s);
                  if i = 0 then
                  MSAAs[5] := s
                  else
                  MSAAs[5] := MSAAs[5] + ', ' + s; }
            end;
        end;
    except
        on E: Exception do
            Showerr(E.Message);
    end;
end;

function TwndMSAAV.IA2Text4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
var
    isp: iserviceprovider;
    iInter: IInterface;
    ia2, ia2Targ: iaccessible2;
    iAV: IAccessibleValue;
    iaTarg: IAccessible;
    hr, i, iRole, t, p, iTarg: integer;
    MSAAs: array [0..11] of widestring;
    rNode, Node: TTreeNode;
    ovChild, ovValue: OleVariant;
    iAL: IAccessibleRelation;
    s, ws: widestring;
    oList, cList: TStringList;



begin
    result := '';
    if not Assigned(iAcc) then Exit;
    if pAcc = nil then pAcc := iAcc;
    if (iMode <> 1) then Exit;
    if (not mnuMSAA.Checked) then Exit;
    //if flgIA2 = 0 then Exit;
    rNode := nil;
    oList := TStringList.Create;
    try

        hr := pAcc.QueryInterface(IID_ISERVICEPROVIDER, isp);
        if SUCCEEDED(hr) then
        begin
            try
            	hr := isp.QueryService(IID_IAccessible, iid_iaccessible2, ia2);
                {if not SUCCEEDED(isp.QueryService(IID_IAccessible, iid_iaccessible2, ia2)) then
                begin
                    hr := E_NOINTERFACE;
                    ia2 := nil;
                end; }
            except
                hr := E_FAIL;
            end;
            if SUCCEEDED(hr) then
            begin
                for i := 0 to 8 do
                begin
                    MSAAs[i] := none;
                end;

                ovChild := varParent;
                if (flgIA2 and TruncPow(2, 0)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accName(ovChild, ws)) then
                        begin
                            MSAAs[0] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[0] := E.Message;
                    end;
                end;

                if (flgIA2 and TruncPow(2, 1)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.role(iRole)) then
                        begin
                            MSAAs[1] := GetIA2Role(iRole);;
                        end;
                    except
                        on E: Exception do
                            MSAAs[1] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 2)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_states(iRole)) then
                        begin
                            MSAAs[2] := GetIA2State(iRole);
                        end;
                    except
                        on E: Exception do
                            MSAAs[2] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 3)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accDescription(ovChild, ws)) then
                        begin
                            MSAAs[3] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[3] := E.Message;
                    end;
                end;



                if (flgIA2 and TruncPow(2, 5)) <> 0 then
                begin

                    cList := TStringList.Create;
                    try

                    try
                        if SUCCEEDED(ia2.Get_attributes(s)) then
                        begin
                            cList.Delimiter := ';';
  													cList.DelimitedText := s;
                            for i := 0 to cList.Count - 1 do
                            begin
                            		if cList[i] = '' then continue;

                                if i = 0 then
                                    MSAAs[5] := cList[i]
                                else
                                    MSAAs[5] := MSAAs[5] + ', ' + cList[i];
                                oList.Add(cList[i]);

                            end;
                        end;
                    except
                        on E: Exception do
                        begin
                            MSAAs[5] := MSAAs[5] + E.Message;
                            oList.Add(E.Message);
                        end;
                    end;
                    finally
                    	cList.Free;
                    end;

                end;

                if (flgIA2 and TruncPow(2, 6)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accValue(ovChild, ws)) then
                        begin
                            MSAAs[6] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[6] := E.Message;
                    end;
                end;



                if (flgIA2 and TruncPow(2, 7)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_localizedExtendedRole(ws)) then
                        begin
                            MSAAs[7] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[7] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 8)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_localizedExtendedStates(10, PWideString1(ws), iRole)) then
                        begin
                            MSAAs[8] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[8] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 9)) <> 0 then
                begin
                    try
                        if SUCCEEDED(isp.QueryService(IID_IAccessible, iid_iaccessiblevalue, iAV)) then
                        begin
                            if SUCCEEDED(iAV.Get_currentValue(ovValue)) then
                            	MSAAs[9] := ovValue;
                            if SUCCEEDED(iAV.Get_minimumValue(ovValue)) then
                            	MSAAs[10] := ovValue;
                            if SUCCEEDED(iAV.Get_maximumValue(ovValue)) then
                            	MSAAs[11] := ovValue;
                        end;
                    except
                        on E: Exception do
                            MSAAs[9] := E.Message;
                    end;
                end;
                Result := Result + #13#10 + tab + #9 + lIA2[0] + '<br>';
                for i := 0 to 9 do
                begin
                    if (flgIA2 and TruncPow(2, i)) <> 0 then
                    begin
                        if i = 5 then
                        begin
                            Result := Result + #13#10 + tab + #9 + lIA2[i+1] + ':' + MSAAs[i];// + '(';// + #13#10;

                            	for p := 0 to oList.Count - 1 do
                              begin
                                Result := Result + oList[p];
                                {if p <>  oList.Count - 1 then
                                	Result := Result + ','; }
                              end;
                              Result := Result + '<br>';//');<br>';
                        end
                        else if (i = 4) {and ((flgIA2 and 16) <> 0)} then
                        begin
                            node := nil;

                            try
                                if SUCCEEDED(ia2.Get_nRelations(iRole)) then
                                begin
                                    for p := 0 to iRole - 1 do
                                    begin
                                        Result := Result + #13#10 + tab + #9 + rType + ':';
                                        ia2.Get_relation(p, iAL);
                                        iAL.Get_RelationType(s);

                                        Result := Result + s + '<br>';
                                        Result := Result + #13#10 + tab + #9 + rTarg + ':';


                                        iAL.Get_nTargets(iTarg);
                                        for t := 0 to iTarg do
                                        begin
                                            iInter := nil;
                                            hr := iAL.Get_target(t, iInter);
                                            if SUCCEEDED(hr) then
                                            begin
                                                hr := iInter.QueryInterface(IID_IACCESSIBLE2, ia2Targ);
                                                if SUCCEEDED(hr) then
                                                begin
                                                    hr := ia2Targ.QueryInterface(IID_ISERVICEPROVIDER, isp);
                                                    if SUCCEEDED(hr) then
                                                    begin
                                                        hr := isp.QueryService(IID_IAccessible, iid_iaccessible, iaTarg);
                                                        if SUCCEEDED(hr) then
                                                        begin

                                                            //s := iaTarg.accName[0];
                                                            iaTarg.Get_accName(0, s);
                                                            if s = '' then
                                                                s := none;
                                                            //iaTarg.Get_accValue(0, s);
                                                            ia2Targ.role(iRole);
                                                            ws := GetIA2Role(iRole);
                                                            if ws = '' then
                                                                ws := none;
                                                            //iaTarg.Get_accValue(0, ws);
                                                            Result := Result + s + ' - ' + ws;
                                                            {if t <> itarg then
                                                            	Result := Result + ',';}
                                                        end;
                                                    end;
                                                end;
                                            end;
                                        end;
                                        Result := Result + ';<br>';
                                    end;
                                end;
                            except
                                on E:Exception do
                                    ShowErr(E.Message);
                            end;
                        end
                        else if i = 9 then
                        begin

                          for t := 0 to 2 do
                          begin
                            Result := Result + #13#10 + tab + #9 + lIA2[11+t] + ':' + MSAAs[t+9] + ';<br>';
                          end;

                        end
                        else
                        begin

                            Result := Result + #13#10 + tab + #9 + lIA2[i+1] + ':' + MSAAs[i] + ';<br>';
                        end;
                    end;
                end;


                //Result := Result + lIA2[9] + ':' + #9 + MSAAs[8] + #13#10;

            end;
        end;
    finally
        oList.Free;
    end;

end;

function TwndMSAAV.SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
var
    isp: iserviceprovider;
    iInter: IInterface;
    ia2, ia2Targ: iaccessible2;
    iAV: IAccessibleValue;
    iaTarg: IAccessible;
    hr, i, iRole, t, p, iTarg: integer;
    MSAAs: array [0..11] of widestring;
    rNode, Node, cNode, ccNode, cccNode: TTreeNode;
    nodes: TTreeNodes;
    ovChild, ovValue: OleVariant;
    iAL: IAccessibleRelation;
    s, ws: widestring;
    oList, cList: TStringList;
    TD: PTreeData;


    
begin
    result := '';
    if not Assigned(iAcc) then Exit;
    if pAcc = nil then pAcc := iAcc;
    if (iMode <> 1) then Exit;
    if (not mnuMSAA.Checked) then Exit;
    //if flgIA2 = 0 then Exit;
    rNode := nil;
    oList := TStringList.Create;
    try

        hr := pAcc.QueryInterface(IID_ISERVICEPROVIDER, isp);
        if SUCCEEDED(hr) then
        begin
            try
            	hr := isp.QueryService(IID_IAccessible, iid_iaccessible2, ia2);
                {if not SUCCEEDED(isp.QueryService(IID_IAccessible, iid_iaccessible2, ia2)) then
                begin
                    hr := E_NOINTERFACE;
                    ia2 := nil;
                end; }
            except
                hr := E_FAIL;
            end;
            if SUCCEEDED(hr) then
            begin
                for i := 0 to 8 do
                begin
                    MSAAs[i] := none;
                end;

                ovChild := varParent;
                if (flgIA2 and TruncPow(2, 0)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accName(ovChild, ws)) then
                        begin
                            MSAAs[0] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[0] := E.Message;
                    end;
                end;

                if (flgIA2 and TruncPow(2, 1)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.role(iRole)) then
                        begin
                            MSAAs[1] := GetIA2Role(iRole);;
                        end;
                    except
                        on E: Exception do
                            MSAAs[1] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 2)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_states(iRole)) then
                        begin
                            MSAAs[2] := GetIA2State(iRole);
                        end;
                    except
                        on E: Exception do
                            MSAAs[2] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 3)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accDescription(ovChild, ws)) then
                        begin
                            MSAAs[3] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[3] := E.Message;
                    end;
                end;



                if (flgIA2 and TruncPow(2, 5)) <> 0 then
                begin

                    cList := TStringList.Create;
                    try

                    try
                        if SUCCEEDED(ia2.Get_attributes(s)) then
                        begin
                            cList.Delimiter := ';';
  													cList.DelimitedText := s;
                            for i := 0 to cList.Count - 1 do
                            begin
                            		if cList[i] = '' then continue;

                                if i = 0 then
                                    MSAAs[5] := cList[i]
                                else
                                    MSAAs[5] := MSAAs[5] + ', ' + cList[i];
                                oList.Add(cList[i]);

                            end;
                        end;
                    except
                        on E: Exception do
                        begin
                            MSAAs[5] := MSAAs[5] + E.Message;
                            oList.Add(E.Message);
                        end;
                    end;
                    finally
                    	cList.Free;
                    end;

                end;

                if (flgIA2 and TruncPow(2, 6)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_accValue(ovChild, ws)) then
                        begin
                            MSAAs[6] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[6] := E.Message;
                    end;
                end;



                if (flgIA2 and TruncPow(2, 7)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_localizedExtendedRole(ws)) then
                        begin
                            MSAAs[7] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[7] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 8)) <> 0 then
                begin
                    try
                        if SUCCEEDED(ia2.Get_localizedExtendedStates(10, PWideString1(ws), iRole)) then
                        begin
                            MSAAs[8] := ws;
                        end;
                    except
                        on E: Exception do
                            MSAAs[8] := E.Message;
                    end;
                end;
                if (flgIA2 and TruncPow(2, 9)) <> 0 then
                begin
                    try
                        if SUCCEEDED(isp.QueryService(IID_IAccessible, iid_iaccessiblevalue, iAV)) then
                        begin
                            if SUCCEEDED(iAV.Get_currentValue(ovValue)) then
                            	MSAAs[9] := ovValue;
                            if SUCCEEDED(iAV.Get_minimumValue(ovValue)) then
                            	MSAAs[10] := ovValue;
                            if SUCCEEDED(iAV.Get_maximumValue(ovValue)) then
                            	MSAAs[11] := ovValue;
                        end;
                    except
                        on E: Exception do
                            MSAAs[9] := E.Message;
                    end;
                end;
                Result := lIA2[0] + #13#10;
                nodes := TreeList1.Items;
                if SetTL then
                    rNode := nodes.AddChild(nil, lIA2[0]);
                for i := 0 to 9 do
                begin
                    if (flgIA2 and TruncPow(2, i)) <> 0 then
                    begin
                        if i = 5 then
                        begin
                            Result := Result + lIA2[i+1] + ':' + #9 + MSAAs[i] + '(';// + #13#10;
                            if SetTL then
                            begin
                                node := nodes.AddChild(rNode, lIA2[i+1]);
                                for p := 0 to oList.Count - 1 do
                                begin
                                    t := pos(':', oList[p]);
                                    cnode := nodes.AddChild(Node, copy(oList[p], 0, t-1));
                                    TreeList1.SetNodeColumn(cnode, 1, copy(oList[p], t+1, Length(oList[p])));
                                end;
                            end
                            else
                            begin
                            	for p := 0 to oList.Count - 1 do
                              begin
                                Result := Result + oList[p];
                              end;
                              Result := Result + ')' + #13#10;
                            end;
                        end
                        else if (i = 4) {and ((flgIA2 and 16) <> 0)} then
                        begin
                           node := nil;

                            try
                                if SUCCEEDED(ia2.Get_nRelations(iRole)) then
                                begin
                                    for p := 0 to iRole - 1 do
                                    begin
                                        Result := Result + #13#10 + #9 + rType + ':';
                                        ia2.Get_relation(p, iAL);
                                        iAL.Get_RelationType(s);




                                        if SetTL then
                                        begin
                                            node := nodes.AddChild(rNode, lIA2[i+1]);
                                            cnode := nodes.AddChild(Node, inttostr(p));
                                            ccNode := nodes.AddChild(cNode, rType);
                                            TreeList1.SetNodeColumn(ccnode, 1, s);
                                            ccNode := nodes.AddChild(cNode, rTarg);
                                        end;
                                        Result := Result + s;
                                        Result := Result + #13#10 + #9 + rTarg + ':';


                                        iAL.Get_nTargets(iTarg);
                                        for t := 0 to iTarg do
                                        begin
                                            iInter := nil;
                                            hr := iAL.Get_target(t, iInter);
                                            if SUCCEEDED(hr) then
                                            begin
                                                hr := iInter.QueryInterface(IID_IACCESSIBLE2, ia2Targ);
                                                if SUCCEEDED(hr) then
                                                begin
                                                    hr := ia2Targ.QueryInterface(IID_ISERVICEPROVIDER, isp);
                                                    if SUCCEEDED(hr) then
                                                    begin
                                                        hr := isp.QueryService(IID_IAccessible, iid_iaccessible, iaTarg);
                                                        if SUCCEEDED(hr) then
                                                        begin

                                                            //s := iaTarg.accName[0];
                                                            iaTarg.Get_accName(0, s);
                                                            if s = '' then
                                                                s := none;
                                                            //iaTarg.Get_accValue(0, s);
                                                            ia2Targ.role(iRole);
                                                            ws := GetIA2Role(iRole);
                                                            if ws = '' then
                                                                ws := none;
                                                            //iaTarg.Get_accValue(0, ws);
                                                            Result := Result + s + ' - ' + ws + #13#10 + #9;
                                                            if SetTL then
                                                            begin
                                                                cccNode := nodes.AddChild(ccNode, inttostr(t));
                                                                TreeList1.SetNodeColumn(cccnode, 1, s + ' - ' + ws);
                                                                New(TD);
                                                                TD^.Acc := iaTarg;
                                                                TD^.iID := 0;
                                                                cccNode.Data := Pointer(TD);
                                                            end;
                                                        end;
                                                    end;
                                                end;
                                            end;
                                        end;

                                    end;
                                end;
                            except
                                on E:Exception do
                                    ShowErr(E.Message);
                            end;
                        end
                        else if i = 9 then
                        begin
                        	if SetTL then
                          begin
                        		node := nodes.AddChild(rNode, lIA2[10]);
                          end;
                          for t := 0 to 2 do
                          begin
                          	if SetTL then
                            begin
                          		cnode := nodes.AddChild(Node, lIA2[11+t]);
                          		TreeList1.SetNodeColumn(cnode, 1, MSAAs[t+9]);
                            end;
                            Result := Result + lIA2[11+t] + ':' + #9 + MSAAs[t+9] + #13#10;
                          end;

                        end
                        else
                        begin
                            if SetTL then
                            begin
                                node := nodes.AddChild(rNode, lIA2[i+1]);
                                TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
                            end;
                            Result := Result + lIA2[i+1] + ':' + #9 + MSAAs[i] + #13#10;
                        end;
                    end;
                end;


                //Result := Result + lIA2[9] + ':' + #9 + MSAAs[8] + #13#10;
                if SetTL then
                    rNode.Expand(True);
            end;
        end;
        TipTextIA2 := Result + #13#10 + #13#10;
    finally
        oList.Free;
    end;


end;

function TwndMSAAV.MSAAText4Tip(pAcc: IAccessible = nil): string;
var
    PC:PChar;
    ovValue, ovChild: OleVariant;
    MSAAs: array [0..10] of widestring;
    ws: widestring;
    i: integer;
    pDis: IDispatch;
    pa: IAccessible;
begin
    if not Assigned(iAcc) then Exit;
    if not mnuMSAA.Checked then Exit;
    if pAcc = nil then pAcc := iAcc;
    for i := 0 to 10 do
    begin
        MSAAs[i] := none;
    end;
    try

        ovChild := varParent;

        if (flgMSAA and 1) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accName(ovChild, ws)) then
                    MSAAs[0] := ws;
            except
                on E: Exception do
                    MSAAs[0] := E.Message;
            end;
        end;
        if (flgMSAA and 2) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accRole(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            PC := StrAlloc(255);
                            GetRoleTextW(ovValue, PC, StrBufSize(PC));
                            MSAAs[1] := PC;
                            StrDispose(PC);
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[1] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[1] := E.Message;
            end;
        end;
        if (flgMSAA and 4) <> 0 then
        begin
             try
                if SUCCEEDED(pAcc.Get_accState(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            MSAAs[2] := GetMultiState(Cardinal(ovValue));
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[2] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[2] := E.Message;
            end;
        end;
        if (flgMSAA and 8) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDescription(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[3] := ws;
            except
                on E: Exception do
                    MSAAs[3] := E.Message;
            end;
        end;

        if (flgMSAA and 16) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDefaultAction(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[4] := ws;
            except
                on E: Exception do
                    MSAAs[4] := E.Message;
            end;
        end;

        if (flgMSAA and 32) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accValue(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[5] := ws;
            except
                on E: Exception do
                    MSAAs[5] := E.Message;
            end;
        end;

        if (flgMSAA and 64) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accParent(pDis)) then
                begin
                    if Assigned(pDis) then
                    begin
                        if SUCCEEDED(pDis.QueryInterface(IID_IACCESSIBLE, pa)) then
                        begin
                            pa.Get_accName(CHILDID_SELF, MSAAs[6]);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[6] := E.Message;
            end;
        end;

        if (flgMSAA and 128) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accChildCount(i)) then
                    MSAAs[7] := IntTostr(i);
            except
                on E: Exception do
                    MSAAs[7] := E.Message;
            end;
        end;

        if (flgMSAA and 256) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelp(ovChild, ws)) then
                    MSAAs[8] := ws;
            except
                on E: Exception do
                    MSAAs[8] := E.Message;
            end;
        end;
        if (flgMSAA and 512) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelpTopic(ws, ovChild, i)) then
                begin
                    MSAAs[9] := ws + ' , ' + InttoStr(i);
                end;
            except
                on E: Exception do
                    MSAAs[9] := E.Message;
            end;
        end;
        if (flgMSAA and 1024) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accKeyboardShortcut(ovChild, ws)) then
                    MSAAs[10] := ws;
            except
                on E: Exception do
                    MSAAs[10] := E.Message;
            end;
        end;

        Result := lMSAA[0] + #13#10;
        for i := 0 to 10 do
        begin
            if (flgMSAA and TruncPow(2, i)) <> 0 then
            begin
                Result := Result + lMSAA[i+1] + ':' + #9 + MSAAs[i] + #13#10;
            end;
        end;
        Result := result + #13#10;
    finally

    end;
end;

function TwndMSAAV.MSAAText4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
var
    PC:PChar;
    ovValue, ovChild: OleVariant;
    MSAAs: array [0..10] of widestring;
    ws: widestring;
    i: integer;
    pDis: IDispatch;
    pa: IAccessible;
begin
    if not Assigned(iAcc) then Exit;
    if not mnuMSAA.Checked then Exit;
    if pAcc = nil then pAcc := iAcc;
    for i := 0 to 10 do
    begin
        MSAAs[i] := none;
    end;
    try

        ovChild := varParent;

        if (flgMSAA and 1) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accName(ovChild, ws)) then
                begin
                    ws := ReplaceStr(ws, '<', '&lt;');
                    ws := ReplaceStr(ws, '>', '&gt;');
                    MSAAs[0] := ws;
                end;
            except
                on E: Exception do
                    MSAAs[0] := E.Message;
            end;
        end;
        if (flgMSAA and 2) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accRole(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            PC := StrAlloc(255);
                            GetRoleTextW(ovValue, PC, StrBufSize(PC));
                            MSAAs[1] := PC;
                            StrDispose(PC);
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[1] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[1] := E.Message;
            end;
        end;
        if (flgMSAA and 4) <> 0 then
        begin
             try
                if SUCCEEDED(pAcc.Get_accState(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            MSAAs[2] := GetMultiState(Cardinal(ovValue));
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[2] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[2] := E.Message;
            end;
        end;
        if (flgMSAA and 8) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDescription(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[3] := ws;
            except
                on E: Exception do
                    MSAAs[3] := E.Message;
            end;
        end;

        if (flgMSAA and 16) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDefaultAction(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[4] := ws;
            except
                on E: Exception do
                    MSAAs[4] := E.Message;
            end;
        end;

        if (flgMSAA and 32) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accValue(ovChild, ws)) then
                if ws <> '' then
                    MSAAs[5] := ws;
            except
                on E: Exception do
                    MSAAs[5] := E.Message;
            end;
        end;

        if (flgMSAA and 64) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accParent(pDis)) then
                begin
                    if Assigned(pDis) then
                    begin
                        if SUCCEEDED(pDis.QueryInterface(IID_IACCESSIBLE, pa)) then
                        begin
                            pa.Get_accName(CHILDID_SELF, MSAAs[6]);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[6] := E.Message;
            end;
        end;

        if (flgMSAA and 128) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accChildCount(i)) then
                    MSAAs[7] := IntTostr(i);
            except
                on E: Exception do
                    MSAAs[7] := E.Message;
            end;
        end;

        if (flgMSAA and 256) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelp(ovChild, ws)) then
                    MSAAs[8] := ws;
                    //MSAAs[8] := pAcc.accHelp[ovChild];
            except
                on E: Exception do
                    MSAAs[8] := E.Message;
            end;
        end;
        if (flgMSAA and 512) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelpTopic(ws, ovChild, i)) then
                begin
                    MSAAs[9] := ws + ' , ' + InttoStr(i);
                end;
            except
                on E: Exception do
                    MSAAs[9] := E.Message;
            end;
        end;
        if (flgMSAA and 1024) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accKeyboardShortcut(ovChild, ws)) then
                    MSAAs[10] := ws;
            except
                on E: Exception do
                    MSAAs[10] := E.Message;
            end;
        end;

        Result := Result + #13#10 + tab + #9 + lMSAA[0] + '<br>';
        for i := 0 to 10 do
        begin
            if (flgMSAA and TruncPow(2, i)) <> 0 then
            begin
                Result := Result + #13#10 + tab + #9 + lMSAA[i+1] + ':' + MSAAs[i] + ';<br>';
            end;
        end;
        //Result := result + #13#10;
    finally

    end;
end;

procedure TwndMSAAV.TreeView1Hint(Sender: TObject; const Node: TTreeNode;
  var Hint: string);
begin
{var
    MSAAs: array[0..2] of string;
    PT: PTreeData;
    pa: IAccessible;
    ini: TMemIniFile;
    ws: widestring;
    ovValue, ovChild: OleVariant;
begin
    if not assigned(Node) then
        Exit;
    if not Extip then
        Exit;
    pa := TTreeData(Node.Data^).Acc;
    if not assigned(pa) then
        Exit;
    ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
    try
        MSAAs[0] := ini.ReadString('MSAA', 'Name', 'Name') + ':';
        MSAAs[1] := ini.ReadString('MSAA', 'State', 'State') + ':';
        MSAAs[2] := ini.ReadString('MSAA', 'Description', 'Description') + ':';

    finally
        ini.Free;
    end;
    ovChild := 0;
    try
        try
            if SUCCEEDED(pa.Get_accName(ovChild, ws)) then
            begin
                if ws = '' then
                    ws := none;
                MSAAs[0] := MSAAs[0] + ws;
            end;
        except
            on E: Exception do
                MSAAs[0] := MSAAs[0] + E.Message;
        end;

        try
            if SUCCEEDED(pa.Get_accState(ovChild, ovValue)) then
            begin
                if VarHaveValue(ovValue) then
                begin
                    if VarIsNumeric(ovValue) then
                    begin
                        MSAAs[1] := MSAAs[1] + GetMultiState(Cardinal(ovValue));
                    end
                    else if VarIsStr(ovValue) then
                    begin
                        MSAAs[1] := MSAAs[1] + VarToStr(ovValue);
                    end;
                end;
            end;
        except
            on E: Exception do
                MSAAs[1] := MSAAs[1] + E.Message;
        end;
        ws := '';
        try
            if SUCCEEDED(pa.Get_accDescription(ovChild, ws)) then
            begin
                if ws = '' then
                    ws := none;
                MSAAs[2] := MSAAs[2] + ws;
            end;

        except
            on E: Exception do
                MSAAs[2] := MSAAs[2] + E.Message;
        end;
        Hint := MSAAs[0] + ' , ' + MSAAs[1] + ' , ' + MSAAs[2];
        //Hint := '';
        //Hint := node.Text + #13#10 + Hint;
    finally

    end;}
end;

function TwndMSAAV.MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = false): string;
var
    PC:PChar;
    ovValue, ovChild: OleVariant;
    MSAAs: array [0..10] of widestring;
    ws: widestring;
    node: TTreeNode;
    nodes: TTreeNodes;
    i: integer;
    pDis: IDispatch;
    pa: IAccessible;
begin
    if not Assigned(iAcc) then Exit;
    if not mnuMSAA.Checked then Exit;
    //if flgMSAA = 0 then Exit;
    if pAcc = nil then pAcc := iAcc;
    for i := 0 to 10 do
    begin
        MSAAs[i] := none;
    end;


        ovChild := varParent;

        if (flgMSAA and 1) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accName(ovChild, ws)) then
                    MSAAs[0] := ws;
            except
                on E: Exception do
                    MSAAs[0] := E.Message;
            end;
        end;
        if (flgMSAA and 2) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accRole(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            PC := StrAlloc(255);
                            GetRoleTextW(ovValue, PC, StrBufSize(PC));
                            MSAAs[1] := PC;
                            StrDispose(PC);
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[1] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[1] := E.Message;
            end;
        end;
        if (flgMSAA and 4) <> 0 then
        begin
             try
                if SUCCEEDED(pAcc.Get_accState(ovChild, ovValue)) then
                begin
                    if VarHaveValue(ovValue) then
                    begin
                        if VarIsNumeric(ovValue) then
                        begin
                            MSAAs[2] := GetMultiState(Cardinal(ovValue));
                        end
                        else if VarIsStr(ovValue) then
                        begin
                            MSAAs[2] := VarToStr(ovValue);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[2] := E.Message;
            end;
        end;
        if (flgMSAA and 8) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDescription(ovChild, ws)) then
                    MSAAs[3] := ws;
            except
                on E: Exception do
                    MSAAs[3] := E.Message;
            end;
        end;

        if (flgMSAA and 16) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accDefaultAction(ovChild, ws)) then
                    MSAAs[4] := ws;
            except
                on E: Exception do
                    MSAAs[4] := E.Message;
            end;
        end;

        if (flgMSAA and 32) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accValue(ovChild, ws)) then
                    MSAAs[5] := ws;
            except
                on E: Exception do
                    MSAAs[5] := E.Message;
            end;
        end;

        if (flgMSAA and 64) <> 0 then
        begin
            try

                if SUCCEEDED(pAcc.Get_accParent(pDis)) then
                begin
                    if Assigned(pDis) then
                    begin
                        if SUCCEEDED(pDis.QueryInterface(IID_IACCESSIBLE, pa)) then
                        begin
                            pa.Get_accName(CHILDID_SELF, MSAAs[6]);
                        end;
                    end;
                end;
            except
                on E: Exception do
                    MSAAs[6] := E.Message;
            end;
        end;

        if (flgMSAA and 128) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accChildCount(i)) then
                    MSAAs[7] := IntTostr(i);
            except
                on E: Exception do
                    MSAAs[7] := E.Message;
            end;
        end;

        if (flgMSAA and 256) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelp(ovChild, ws)) then
                    MSAAs[8] := ws;
            except

                on E: Exception do
                    MSAAs[8] := E.Message;
            end;
        end;
        if (flgMSAA and 512) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accHelpTopic(ws, ovChild, i)) then
                if ws <> '' then
                begin
                    MSAAs[9] := ws + ' , ' + InttoStr(i);
                end;
            except
                on E: Exception do
                    MSAAs[9] := E.Message;
            end;
        end;
        if (flgMSAA and 1024) <> 0 then
        begin
            try
                if SUCCEEDED(pAcc.Get_accKeyboardShortcut(ovChild, ws)) then
                    MSAAs[10] := ws;
            except

                on E: Exception do
                    MSAAs[10] := E.Message;
            end;
        end;

        Result := lMSAA[0] + #13#10;
        if not TextOnly then
        begin
        	nodes := TreeList1.Items;
        	rNode := nodes.AddChild(nil, lMSAA[0]);
        	for i := 0 to 10 do
        	begin
            	if (flgMSAA and TruncPow(2, i)) <> 0 then
            	begin
                	node := nodes.AddChild(rNode, lMSAA[i+1]);
                	TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
                	Result := Result + lMSAA[i+1] + ':' + #9 + MSAAs[i] + #13#10;
                	if i = 4 then
                	begin
                    	iDefIndex := Node.Index;
                	end;
            	end;
        	end;
          rNode.Expand(True);
        end
        else
        begin
         	for i := 0 to 10 do
        	begin
            	if (flgMSAA and TruncPow(2, i)) <> 0 then
            	begin
                	Result := Result + lMSAA[i+1] + ':' + #9 + MSAAs[i] + #13#10;
            	end;
        	end;
        end;
        TipText := result + #13#10;


end;

function TwndMSAAV.HTMLText: string;
var
    HTMLs: array [0..2, 0..1] of widestring;
    iAttrCol: IHTMLATTRIBUTECOLLECTION;
    iDomAttr: IHTMLDOMATTRIBUTE;
    s, sHTML, sType: string;
    hAttrs: widestring;
    ovValue: OleVariant;
    i: integer;
    rNode, Node, sNode: TTreeNode;
    ini: TMemIniFile;

    aPC, aPC2: array [0..64] of pWidechar;
    aSI: array [0..64] of Smallint;
    PC, PC2:PChar;
    SI: Smallint;
    PU, PU2: PUINT;
    WD: Word;
begin
    if ((CEle = nil) and (Imode = 0)) or ((SDOM = nil) and (iMode = 1)) then
    begin
        GetNaviState(True);
        Exit;
    end;
    sHTML := 'HTML';
    HTMLs[0, 0] := 'Element name';
    HTMLs[1, 0] := 'Attributes';
    HTMLs[2, 0] := 'Code';
    if FileExists(Transpath) then
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	    try
            sHTML := ini.ReadString('HTML', 'HTML', sHTML);
            HTMLs[0, 0] := ini.ReadString('HTML', 'Element_name', HTMLs[0, 0]);
            HTMLs[1, 0] := ini.ReadString('HTML', 'Attributes', HTMLs[1, 0]);
            HTMLs[2, 0] := ini.ReadString('HTML', 'Code', HTMLs[2, 0]);
            sType := '(' + ini.ReadString('HTML', 'outerHTML', 'outerHTML') + ')';
            HTMLs[2, 0] := HTMLs[2, 0] + Stype;
        finally
            ini.Free;
        end;
    end;




    if iMode = 0 then
    begin
        HTMLS[0, 1] := CEle.tagName;
        rNode := TreeList1.Items.AddChild(nil, sHTML);
        Node := TreeList1.Items.AddChild(rNode, HTMLS[0, 0]);
        TreeList1.SetNodeColumn(node, 1, HTMLS[0, 1]);
        iAttrCol := (CEle as IHTMLDOMNODE).attributes as IHTMLATTRIBUTECOLLECTION;
        sNode := nil;
        for i := 0 to iAttrCol.length - 1 do
        begin
            ovValue := i;
            iDomAttr := iAttrCol.item(ovValue) as IHTMLDomAttribute;
            if iDomAttr.specified then
            begin
                s := lowerCase(VarToStr(iDomAttr.nodeName));
                if (s <> 'role') and (Copy(s, 1, 4) <> 'aria') then
                begin
                    try
                        if sNode = nil then
                        begin
                            Node := TreeList1.Items.AddChild(rNode, HTMLS[1, 0]);
                            sNode := Node;
                        end;
                        Node := TreeList1.Items.AddChild(sNode, VarToStr(iDOmAttr.nodeName));
                        if VarHaveValue(iDomAttr.nodeValue) then
                        begin
                            TreeList1.SetNodeColumn(node, 1, VarToStr(iDomAttr.nodeValue));
                            hattrs := hattrs + VarToStr(iDOmAttr.nodeName)  + VarToStr(iDomAttr.nodeValue) + '"' + #13#10;
                        end;
                    except
                        on E:Exception do
                        begin
                            ShowErr(E.Message);
                        end;
                    end;
                end;
            end;
        end;
    end
    else if iMode = 1 then
    begin
        SDOM.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);

        HTMLS[0, 1] := PC;
        rNode := TreeList1.Items.AddChild(nil, sHTML);
        Node := TreeList1.Items.AddChild(rNode, HTMLS[0, 0]);
        TreeList1.SetNodeColumn(node, 1, HTMLS[0, 1]);
        SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
        if WD > 0 then
        begin
            Node := TreeList1.Items.AddChild(rNode, HTMLS[1, 0]);
            sNode := Node;

            for i := 0 to WD - 1 do
            begin
                s := LowerCase(aPC[i]);
                if s = 'role' then
                begin
                    //
                end
                else if Copy(s, 1, 4) = 'aria' then
                begin
                    //
                end
                else
                begin
                    try

                        Node := TreeList1.Items.AddChild(sNode, aPC[i]);
                        TreeList1.SetNodeColumn(node, 1, aPC2[i]);
                        hattrs := hattrs + WideString(aPC[i]) + '","' + WideString(aPC2[i]) + #13#10;
                    except

                    end;
                end;
            end;
        end
        else
        begin
            Node := TreeList1.Items.AddChild(rNode, HTMLs[1, 0]);
            TreeList1.SetNodeColumn(node, 1, none)
        end;
    end
    else
        Exit;
    if hattrs = '' then
    begin
        Node := TreeList1.Items.AddChild(rNode, HTMLs[1, 0]);
        TreeList1.SetNodeColumn(node, 1, none)
    end;
    if hattrs = '' then
        HTMLS[1, 1] := none
    else
        HTMLS[1, 1] := hattrs;

    HTMLs[2, 1] := CEle.outerHTML;


        //Node := TreeList1.Items.AddChild(rNode, HTMLS[2, 0]);
        //TreeList1.SetNodeColumn(node, 1, HTMLS[2, 1]);
        Result := sHTML + #13#10 + HTMLS[0, 0] + ':' + #9 +  HTMLS[0, 1] +
                  #13#10 + HTMLS[1, 0] + ':' + #9 +  HTMLS[1, 1] +
                  #13#10 + HTMLS[2, 0] + ':' + #9 +  HTMLS[2, 1] + #13#10#13#10;
        Memo1.Text := HTMLS[2, 0] + ':' + #13#10 +  HTMLS[2, 1] + #13#10#13#10;
        rNode.Expand(True);

end;

function TwndMSAAV.HTMLText4FF: string;
var
    HTMLs: array [0..2, 0..1] of string;
    hAttrs, s, Path, sHTML, sType, sTxt: string;
    i: integer;
    aPC, aPC2: array [0..64] of pWidechar;
    aSI: array [0..64] of Smallint;
    PC, PC2:PChar;
    SI: Smallint;
    PU, PU2: PUINT;
    WD, tWD: Word;
    rNode, Node, sNode: TTreeNode;
    ini: TMemIniFile;
    iText: ISimpleDOMText;
    iSP: iServiceProvider;
begin
    if SDOM = nil then
    begin
        GetNaviState(True);
        Exit;
    end;

    sHTML := 'HTML';
    HTMLs[0, 0] := 'Element name';
    HTMLs[1, 0] := 'Attributes';
    HTMLs[2, 0] := 'Code';
    if FileExists(Transpath) then
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	    try
            sHTML := ini.ReadString('HTML', 'HTML', sHTML);
            HTMLs[0, 0] := ini.ReadString('HTML', 'Element_name', HTMLs[0, 0]);
            HTMLs[1, 0] := ini.ReadString('HTML', 'Attributes', HTMLs[1, 0]);
            HTMLs[2, 0] := ini.ReadString('HTML', 'Code', HTMLs[2, 0]);
            sTxt := ini.ReadString('HTML', 'Text', 'Text');
            sType := '(' + ini.ReadString('HTML', 'innerHTML', 'innerHTML') + ')';
            //HTMLs[2, 0] := HTMLs[2, 0] + Stype;
        finally
            ini.Free;
        end;
    end;
    try
        Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
        SDOM.get_nodeInfo(PC, SI, PC2, PU, PU2, tWD);

        if tWD <> 3 then
            HTMLS[0, 1] := PC
        else
            HTMLS[0, 1] := sTxt;
        rNode := TreeList1.Items.AddChild(nil, sHTML);
        Node := TreeList1.Items.AddChild(rNode, HTMLS[0, 0]);
        TreeList1.SetNodeColumn(node, 1, HTMLS[0, 1]);
        SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
        if WD > 0 then
        begin
            Node := TreeList1.Items.AddChild(rNode, HTMLS[1, 0]);
            sNode := Node;

            for i := 0 to WD - 1 do
            begin
                s := LowerCase(aPC[i]);
                if s = 'role' then
                begin
                    //
                end
                else if Copy(s, 1, 4) = 'aria' then
                begin
                    //
                end
                else
                begin
                    try

                        Node := TreeList1.Items.AddChild(sNode, aPC[i]);
                        TreeList1.SetNodeColumn(node, 1, aPC2[i]);
                        hattrs := hattrs + WideString(aPC[i]) + '","' + WideString(aPC2[i]) + #13#10;
                    except

                    end;
                end;
            end;
        end
        else
        begin
            Node := TreeList1.Items.AddChild(rNode, HTMLs[1, 0]);
            TreeList1.SetNodeColumn(node, 1, none)
        end;
        if hattrs = '' then
            HTMLS[1, 1] := none
        else
            HTMLS[1, 1] := hattrs;
        PC := '';
        if tWD <> 3 then
        begin
            SDOM.get_innerHTML(PC);
            HTMLs[2, 0] := HTMLs[2, 0] + Stype;
        end
        else
        begin
            iSP := SDOM as IServiceProvider;
            if SUCCEEDED(iSP.QueryInterface(IID_ISIMPLEDOMTEXT, iText)) then
            begin
                iText.get_domText(PC);
                HTMLs[2, 0] := HTMLs[2, 0] + '(' + sTxt + ')';
            end;
        end;

        HTMLs[2, 1] := PC;

        Result := sHTML + #13#10 + HTMLS[0, 0] + ':' + #9 +  HTMLS[0, 1] +
                  #13#10 + HTMLS[1, 0] + ':' + #9 +  HTMLS[1, 1] +
                  #13#10 + HTMLS[2, 0] + ':' + #9 +  HTMLS[2, 1] + #13#10#13#10;
        Memo1.Text := HTMLS[2, 0] + ':' + #13#10 +  HTMLS[2, 1] + #13#10#13#10;
        rNode.Expand(True);
    except

    end;
end;

procedure Reflex(sID: string; iChild: integer; ISEle:ISimpleDOMNODE; var outEle: ISimpleDOMNODE);
var
    aPC, aPC2: array [0..1] of pchar;
    aSI: array [0..1] of Smallint;
    i: integer;
    ChildNode: ISimpleDOMNode;
    PC, PC2:PChar;
    SI: Smallint;
    pChild, PU2: PUINT;
    WD: Word;
begin
    aPC[0] := 'id';
    aPC[1] := 'name';
    isEle.get_attributesForNames(2, aPC[0], aSI[0], aPC2[0]);

    if aPC2[0] = sID then
    begin
        //showmessage(sID + '/' + aPC2[0]);
        outEle := isEle;
        exit;
    end;

    if iChild > 0 then
    begin
        for i := 0 to iChild - 1 do
        begin
            if outEle <> nil then
                Break;
            isEle.get_childAt(i, ChildNode);
            ChildNode.get_nodeInfo(PC, SI, PC2, pChild, PU2, WD);
            Reflex(sID, Integer(pChild), ChildNode, outEle);
        end;
    end;
end;

function TwndMSAAV.GetSameAcc: boolean;
var
    i: integer;
    RC1, RC2: TRect;
    rNode: TTreeNode;
const
    IID_IUnknown : TGUID = '{00000000-0000-0000-C000-000000000046}';
begin
    Result := false;
    try
        //if not SUCCEEDED(iAcc.QueryInterface(IID_IUnknown, pUnk1)) then
            //Exit;
        if TBList.Count > 0 then
        begin
            iAcc.accLocation(RC1.Left, RC1.Top, RC1.Right, RC1.Bottom, varParent);
            for i := 0 to TBList.Count - 1 do
            begin
                //Application.ProcessMessages;
                if i > TreeView1.Items.Count - 1 then
                    Break;
                rNOde := TreeView1.Items.GetNode(HTreeItem(TBList.Items[i]));
                if rNode <> nil then
                begin
                    if (rNode.Data <> nil) and (TTreeData(rNode.Data^).Acc <> nil) then
                    begin
                        if AccIsNull(TTreeData(rNode.Data^).Acc) then
                            Break;

                        (TTreeData(rNode.Data^).Acc.accLocation(RC2.Left, RC2.Top, RC2.Right, RC2.Bottom, TTreeData(rNode.Data^).iID));
                        if (RC1.Left = RC2.Left) and (RC1.Top = RC2.Top) and (RC1.Right = RC2.Right) and (RC1.Bottom = RC2.Bottom) then
                        begin
                            try

                                Result := True;
                                TreeView1.OnChange := nil;
                                TreeView1.SetFocus;
                                TreeView1.TopItem := rNode;
                                rNode.Expanded := true;
                                rNode.Selected := True;
                                //GetNaviState;
                                if (acShowTip.Checked) then
                                begin
                                    ShowTipWnd;
                                end;

                                Break;
                            except

                            end;
                        end;
                        //end;
                    end;
                end;
            end;
        end;

    finally
        TreeView1.OnChange := TreeView1Change;
    end;
end;

procedure TwndMSAAV.ReflexID(sID: string; isEle: ISimpleDOMNODE; Labelled: boolean = true);
var
    pEle, TempEle, IDEle: ISimpleDOMNode;
    Serv: IServiceProvider;
    PC, PC2:PChar;
    SI: Smallint;
    iChild, PU2: PUINT;
    WD: Word;
    i: integer;
    lAcc: IAccessible;
    RC: TRect;

begin
    isEle.get_nodeInfo(PC, SI, PC2, iChild, PU2, WD);
    TempEle := isEle;
    i := 0;
    while (LowerCase(String(PC)) <> '#document') do
    begin
        TempEle.get_parentNode(pEle);
        pEle.get_nodeInfo(PC, SI, PC2, iChild, PU2, WD);
        TempEle := pEle;

        if String(PC) = '#document' then
        begin
            iSEle := pEle;
            //showmessage(PC);
            break;
        end;
        inc(i);
        if i > 500 then
        begin
            iSEle := nil;
            Break;
        end;

    end;
    if iSEle <> nil then
    begin
        IDEle := nil;

        //function get_attributesForNames(numAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR): HRESULT; stdcall;
        Reflex(sID, Integer(iChild), ISEle, IDEle);
        //showmessage('E');
    end;
    if IDEle <> nil then
    begin
        IDEle.get_nodeInfo(PC, SI, PC2, iChild, PU2, WD);
        if SUCCEEDED(IDEle.QueryInterface(IID_IServiceProvider, Serv)) then
        begin
            if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, lAcc)) then
            begin
                //if SUCCEEDED(lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0)) then
                //begin
                    lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0);
                    if Labelled then
                        ShowLabeledWnd(clBlue, RC)
                    else
                        ShowDescWnd(clBlue, RC);
                //end;

            end;
        end;
        //showmessage(PC);
    end;
end;

function TwndMSAAV.ARIAText: string;
var
    ARIAs: array [0..1, 0..1] of string;
    iAttrCol: IHTMLATTRIBUTECOLLECTION;
    iDomAttr: IHTMLDOMATTRIBUTE;
    aEle: IHTMLELEMENT;
    s,  sAria: string;
    List, List2: TStringList;
    ovValue: OleVariant;
    i: integer;
    ini: TMemInifile;
    aAttrs, LowerS: string;
    rNode, Node, rcNode: TTreeNode;
    aPC, aPC2: array [0..64] of pchar;
    aSI: array [0..64] of Smallint;
    WD: Word;
    Serv: IServiceProvider;
    lAcc: IAccessible;
    RC: TRect;
begin

    if ((iMode = 0) and (CEle = nil)) or ((iMode =1) and (SDOM = nil)) then
    begin
        GetNaviState(True);
        Exit;
    end;
    if Assigned(WndLabel) then
    begin
        FreeAndNil(Wndlabel);
        Wndlabel := nil;
    end;
    if Assigned(WndDesc) then
    begin
        FreeAndNil(WndDesc);
        WndDesc := nil;
    end;
    if Assigned(WndTarg) then
    begin
        FreeAndNil(WndTarg);
        WndTarg := nil;
    end;

    sAria := 'ARIA';
    ARIAs[0, 0] := 'Role';
    ARIAs[1, 0] := 'Attributes';
    if FileExists(Transpath) then
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	    try
            sAria := ini.ReadString('ARIA', 'ARIA', sAria);
            ARIAs[0, 0] := ini.ReadString('ARIA', 'Role', ARIAs[0, 0]);
            ARIAs[1, 0] := ini.ReadString('ARIA', 'Attributes', ARIAs[1, 0]);
        finally
            ini.Free;
        end;
    end;
        if (iMode = 0) then
        begin
            iAttrCol := (CEle as IHTMLDOMNODE).attributes as IHTMLATTRIBUTECOLLECTION;
            for i := 0 to iAttrCol.length - 1 do
            begin
                ovValue := i;
                iDomAttr := iAttrCol.item(ovValue) as IHTMLDomAttribute;
                if iDomAttr.specified then
                begin
                    s := lowerCase(VarToStr(iDomAttr.nodeName));
                    if s = 'role' then
                    begin
                        try
                            if VarHaveValue(iDomAttr.nodeValue) then
                                Arias[0, 1] := VarToStr(iDomAttr.nodeValue)
                            else
                                Arias[0, 1] := ''
                        except
                            on E:Exception do
                            begin
                                ShowErr(E.Message);
                                Arias[0, 1] := ''
                            end;
                        end;
                    end
                    else if Copy(s, 1, 4) = 'aria' then
                    begin

                        try
                            if VarHaveValue(iDomAttr.nodeValue) then
                            begin

                                aattrs := aattrs + '"' + VarToStr(iDOmAttr.nodeName) + '","' + (iDomAttr.nodeValue) + '"' + #13#10;
                                LowerS := LowerCase(VarToStr(iDOmAttr.nodeName));
                                if acRect.Checked then
                                begin
                                    if LowerS = 'aria-labelledby' then
                                    begin
                                        aEle := (CEle.document as IHTMLDOCUMENT3).getElementById(VarToStr(iDOmAttr.nodeValue));
                                        if ASsigned(aEle) then
                                        begin
                                            if SUCCEEDED(aEle.QueryInterface(IID_IServiceProvider, Serv)) then
                                            begin
                                                if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, lAcc)) then
                                                begin
                                                    //if SUCCEEDED(lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0)) then
                                                    //begin
                                                        lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0);
                                                        ShowLabeledWnd(clBlue, RC);
                                                    //end;

                                                end;
                                            end;
                                        end;
                                    end
                                    else if LowerS = 'aria-decirbedby' then
                                    begin
                                        aEle := (CEle.document as IHTMLDOCUMENT3).getElementById(VarToStr(iDOmAttr.nodeValue));
                                        if ASsigned(aEle) then
                                        begin
                                            if SUCCEEDED(aEle.QueryInterface(IID_IServiceProvider, Serv)) then
                                            begin
                                                if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, lAcc)) then
                                                begin
                                                    //if SUCCEEDED(lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0)) then
                                                    //begin
                                                        lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0);
                                                        ShowDescWnd(clBlue, RC);
                                                    //end;
                                                end;

                                            end;
                                        end;
                                    end;
                                end;
                            end;

                        except
                            on E:Exception do
                            begin
                                ShowErr(E.Message);
                                aattrs := aattrs + '';
                            end;
                        end;
                    end;
                end;
            end;
        end
        else if (iMode =1) then
        begin
            //function get_attributes(maxAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR; out numAttribs: WORD): HRESULT; stdcall;
            //function get_attributesForNames(numAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR): HRESULT; stdcall;
            SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
            for i := 0 to WD - 1 do
            begin
                LowerS := LowerCase(aPC[i]);
                if LowerS = 'role' then
                begin
                    Arias[0, 1] := aPC2[i];
                end
                else if Copy(LowerS, 1, 4) = 'aria' then
                begin

                    aattrs := aattrs + '"' +  aPC[i] + '","' + aPC2[i] + '"' +  #13#10;
                    if acRect.Checked then
                    begin
                        if LowerS = 'aria-labelledby' then
                        begin
                            //showmessage(aPC2[i]);
                            ReflexID(aPC2[i], SDOM, True);
                        end
                        else if LowerS = 'aria-decirbedby' then
                            ReflexID(aPC2[i], SDOM, False);
                    end;
                end;
            end;
        end;
        if ARIAs[0, 1] = '' then
            ARIAs[0, 1] := none;
        if aattrs = '' then
            ARIAS[1, 1] := none
        else
            ARIAs[1, 1] := aattrs;

        rNode := TreeList1.Items.AddChild(nil, sAria);
        Node := TreeList1.Items.AddChild(rNode, Arias[0, 0]);
        TreeList1.SetNodeColumn(node, 1, Arias[0, 1]);
        Node := TreeList1.Items.AddChild(rNode, Arias[1, 0]);
        if aattrs = '' then
            TreeList1.SetNodeColumn(node, 1, Arias[1, 1])
        else
        begin
            rcNode := Node;
            List := TStringList.Create;
            List2 := TStringList.Create;
            try
                List.Text := aattrs;
                for i := 0 to List.Count - 1 do
                begin
                    List2.Clear;
                    List2.CommaText := List[i];
                    if List2.Count >= 1 then
                    begin
                        Node := TreeList1.Items.AddChild(rcNode, List2[0]);
                        if List2.Count >= 2 then
                        begin
                            TreeList1.SetNodeColumn(node, 1, List2[1]);
                        end;
                    end;
                end;
            finally
                List.Free;
                List2.Free;
            end;
        end;
        Result := sAria + #13#10 + Arias[0, 0] + ':' + #9 +  Arias[0, 1] +
                  #13#10 + Arias[1, 0] + ':' + #9 +  Arias[1, 1] + #13#10#13#10;
        rNode.Expand(True);
end;

function TwndMSAAV.IsSameUIElement(ia1, ia2: IAccessible): boolean;
var
    UIEle1, UIEle2: IUIAutomationElement;
    iSame, iRes: integer;
begin
    result := False;
    if Assigned(UIAuto) and Assigned(ia1) and Assigned(ia2) then
    begin
        if SUCCEEDED(UIAuto.ElementFromIAccessible(ia1, 0, UIEle1)) then
        begin
            if SUCCEEDED(UIAuto.ElementFromIAccessible(ia2, 0, UIEle2)) then
            begin
                iRes := UIAuto.CompareElements(UIEle1, UIEle2, iSame);
                if SUCCEEDED(iRes) and (iSame <> 0) then
                    Result := True;
            end;
        end;

    end;
end;

function TwndMSAAV.UIAText: string;
var
    iRes, i, t, iLen, iBool: integer;
    dblRV: Double;
    rNode, Node, cNode: TTreeNode;
    nodes: TTreeNodes;
    ini: TMemInifile;
    ResEle: IUIAutomationElement;
    UIArray: IUIAutomationElementArray;
    iaTarg: IAccessible;
    oVal: OleVariant;
    ws: widestring;
    RC:  tagRect;
    p:pointer;
    OT: OrientationType;
    iInt: IInterface;
    UIAs: array [0..58] of string;
    TD: PTreeData;
    iLeg: IUIAutomationLegacyIAccessiblePattern;
    iRV: IUIAutomationRangeValuePattern;
    iIface: IInterface;
    hr: hresult;
const
		IsProps: Array [0..19] of Integer
	    = (30027, 30028, 30029, 30030, 30031, 30032, 30033, 30034, 30035, 30036, 30037, 30038, 30039, 30040,
      30041, 30042, 30043, 30044, 30108, 30109);


    function GetMSAA: IAccessible;
    var
        rIacc: IAccessible;
    begin
        Result := nil;
        if SUCCEEDED(ResEle.GetCurrentPattern(10018, iInt)) then
        begin
            if SUCCEEDED(iInt.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern, iLEg)) then
            begin
                if SUCCEEDED(iLeg.GetIAccessible(rIacc)) then
                    Result := rIacc;
            end;
        end;

    end;
    function PropertyPatternIS(iID: integer): string;
    begin

        case iID of
            30001: Result := 'rect';
            30002, 30012: Result := 'int';
            30003: Result := 'controltypeid';
            30008, 30009, 30010, 30016, 30017, 30019, 30022, 30025, 30103: Result := 'bool';
            //30018: Result := 'iuiautomationelement';
            30020: Result := 'int';//'uia_hwnd';
            //30104, 30105, 30106: Result := 'iuiautomationelementarray';
            30023: Result := 'orientationtype';
            else result := 'str';
        end;
    end;
    function GetCID(iid: integer):string;
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
        try
            result := ini.ReadString('UIA', inttoStr(iid), none);
        finally
            ini.Free;
        end;
    end;
    function GetOT:string;
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
        try
            if OT = OrientationType_Horizontal then
                result := ini.ReadString('UIA', 'OrientationType_Horizontal', 'Horizontal')
            else if OT = OrientationType_Vertical then
                result := ini.ReadString('UIA', 'OrientationType_Vertical', 'Vertical')
            else
                result := ini.ReadString('UIA', 'OrientationType_None', 'None');
        finally
            ini.Free;
        end;
    end;

begin
    if not Assigned(iAcc) then Exit;
    if UIAuto = nil then Exit;
    if UIEle = nil then Exit;
    //if flgUIA = 0 then Exit;
    //UIEle := nil;

    try
        //iRes := UIA.ElementFromIAccessible(iAcc, 0, UIEle);
        if {(iRes = S_OK) and} (Assigned(UIEle)) then
        begin
            for i := 0 to 32 do
                UIAs[i] := none;

            if (flgUIA and TruncPow(2, 0)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentAcceleratorKey(ws)) then
                    begin
                        UIAs[0] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[0] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 1)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentAccessKey(ws)) then
                    begin
                        UIAs[1] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[1] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 2)) <> 0 then
            begin
                try

                    if SUCCEEDED(UIEle.Get_CurrentAriaProperties(ws)) then
                    begin
                        UIAs[2] := ws;
                    end;


                except
                    on E: Exception do
                        UIAs[2] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 3)) <> 0 then
            begin
                try
                    VarClear(oVal);
                    TVarData(oVal).VType := varString;
                    UIEle.Get_CurrentAriaRole(ws);
                    if ws <> '' then
                        UIAs[3] := ws;
                except
                    on E: Exception do
                        UIAs[3] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 4)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentAutomationId(ws)) then
                    begin
                        UIAs[4] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[4] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 5)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentBoundingRectangle(RC)) then
                    begin
                        UIAs[5] := IntToStr(RC.Left) + ',' + IntToStr(RC.Top) + ',' + IntToStr(RC.Right) + ',' + IntToStr(RC.Bottom);
                    end;
                except
                    on E: Exception do
                        UIAs[5] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 6)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentClassName(ws)) then
                    begin
                        UIAs[6] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[6] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 7)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentControlType(i)) then
                    begin
                        UIAs[7] := GetCID(i);
                    end;
                except
                    on E: Exception do
                        UIAs[7] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 8)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentCulture(i)) then
                    begin
                        UIAs[8] := InttoStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[8] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 9)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentFrameworkId(ws)) then
                    begin
                        UIAs[9] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[9] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 10)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentHasKeyboardFocus(i)) then
                    begin
                        UIAs[10] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[10] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 11)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentHelpText(ws)) then
                    begin
                        UIAs[11] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[11] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 12)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsControlElement(i)) then
                    begin
                        UIAs[12] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[12] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 13)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsContentElement(i)) then
                    begin
                        UIAs[13] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[13] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 14)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsDataValidForForm(i)) then
                    begin
                        UIAs[14] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[14] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 15)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsEnabled(i)) then
                    begin
                        UIAs[15] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[15] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 16)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsKeyboardFocusable(i)) then
                    begin
                        UIAs[16] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[16] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 17)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsOffscreen(i)) then
                    begin
                        UIAs[17] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[17] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 18)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsPassword(i)) then
                    begin
                        UIAs[18] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[18] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 19)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentIsRequiredForForm(i)) then
                    begin
                        UIAs[19] := IntToBoolStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[19] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 20)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentItemStatus(ws)) then
                    begin
                        UIAs[20] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[20] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 21)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentItemType(ws)) then
                    begin
                        UIAs[21] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[21] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 22)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentLocalizedControlType(ws)) then
                    begin
                        UIAs[22] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[22] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 23)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentName(ws)) then
                    begin
                        UIAs[23] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[23] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 24)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentNativeWindowHandle(p)) then
                    begin
                        UIAs[24] := InttoStr(Integer(p));
                    end;
                except
                    on E: Exception do
                        UIAs[24] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 25)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentOrientation(OT)) then
                    begin

                        UIAs[25] := GetOT;
                    end;
                except
                    on E: Exception do
                        UIAs[25] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 26)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentProcessId(i)) then
                    begin
                        UIAs[26] := InttoStr(i);
                    end;
                except
                    on E: Exception do
                        UIAs[26] := E.Message;
                end;
            end;
            if (flgUIA and TruncPow(2, 27)) <> 0 then
            begin
                try
                    if SUCCEEDED(UIEle.Get_CurrentProviderDescription(ws)) then
                    begin
                        UIAs[27] := ws;
                    end;
                except
                    on E: Exception do
                        UIAs[27] := E.Message;
                end;
            end;



            if (flgUIA2 and TruncPow(2, 1)) <> 0 then
            begin
                try //LiveSetting
                    if SUCCEEDED(UIEle.GetCurrentPropertyValue(30135, oval)) then
                    begin
                        if VarHaveValue(oval) then
                        begin
                            if VarIsStr(oval) or VarIsNumeric(oval) then
                                UIAs[32] := oval;
                        end;
                    end;
                except
                    on E: Exception do
                        UIAs[32] := E.Message;
                end;
            end;

            for i := 0 to 19 do
            begin
            	if (flgUIA2 and TruncPow(2, i + 2)) <> 0 then
            	begin
              	try //LiveSetting
                    if SUCCEEDED(UIEle.GetCurrentPropertyValue(IsProps[i], oval)) then
                    begin
                        if VarHaveValue(oval) then
                        begin
                            if VarIsType(oval, varBoolean)  then
                            begin
                                iBool := oval;
                                UIAs[33+i] := IfThen(iBool <> 0, sTrue, sFalse);
                            end;
                        end;
                    end;
                except
                    on E: Exception do
                        UIAs[33+i] := E.Message;
                end;
              end;
            end;
            if (flgUIA2 and TruncPow(2, 22)) <> 0 then
            begin
            	//#30047~30052
              //GetCurrentPatternAs http://msdn.microsoft.com/en-us/library/windows/desktop/ee696039(v=vs.85).aspx
              //UIA_RangeValuePatternId 10003
              try
              hr :=  UIEle.GetCurrentPattern(UIA_RangeValuePatternId, iiFace);
              if hr = S_OK then
              begin
              	if Assigned(iiFace) then
                begin
                	if SUCCEEDED(iiFace.QueryInterface(IID_IUIAutomationRangeValuePattern, iRV)) then
                  begin
                  	if SUCCEEDED(iRV.Get_CurrentValue(dblRV)) then
                  		UIAs[53] := Floattostr(dblRV);
                    if SUCCEEDED(iRV.Get_CurrentIsReadOnly(iBool)) then
                  		UIAs[54] := IfThen(iBool <> 0, sTrue, sFalse);
                    if SUCCEEDED(iRV.Get_CurrentMinimum(dblRV)) then
                  		UIAs[55] := Floattostr(dblRV);
                    if SUCCEEDED(iRV.Get_CurrentMaximum(dblRV)) then
                  		UIAs[56] := Floattostr(dblRV);
                    if SUCCEEDED(iRV.Get_CurrentLargeChange(dblRV)) then
                  		UIAs[57] := Floattostr(dblRV);
                    if SUCCEEDED(iRV.Get_CurrentSmallChange(dblRV)) then
                  		UIAs[58] := Floattostr(dblRV);
                  end;
                end;
              end;
              except
              	    on E: Exception do
                        UIAs[53] := E.Message;
              end;
              {try
              hr := UIEle.GetCurrentPatternAs(UIA_RangeValuePatternId,  IID_IUIAutomationRangeValuePattern, pRV);
              //hr := UIEle.GetCurrentPatternAs(UIA_RangeValuePatternId,  IID_IRangeValueProvider, pRv);
              iRV := IUIAutomationRangeValuePattern(pRV^);
              //hr := iRV.SetValue(1.0); //iRV.Get_CurrentValue(dblRV);
              hr := iRVprov.Get_Value(dblRV);
              caption := inttostr(hr);
              except
              	    on E: Exception do
                        caption := E.Message;
              end;}
              {if SUCCEEDED(UIEle.GetCurrentPatternAs(UIA_RangeValuePatternId, IID_IUIAutomationRangeValuePattern, pointer(iRV))) then
              begin
              	if SUCCEEDED(iRV.Get_CurrentValue(dblRV)) then
                 	//UIAs[53] := 'PK';//Floattostr(dblRV);
              end;  }
              {for i := 0 to 5 do
              begin
              	try
              		if SUCCEEDED(UIEle.GetCurrentPropertyValue(30047+i, oval)) then
                	begin
                  	if VarHaveValue(oval) then
                  	begin
                    	if VarIsType(oval, varDouble)  then
                  		begin
                  			UIAs[53+i] := Floattostr(oval);
                  		end
                      else if VarIsType(oval, varBoolean)  then
                      begin
                      	iBool := oval;
                        UIAs[53+i] := IfThen(iBool <> 0, sTrue, sFalse);
                      end;
                  	end;
                	end;
                except
                    on E: Exception do
                        UIAs[53+i] := E.Message;
                end;
              end; }
            end;
            Result := lUIA[0] + #13#10;
            nodes := TreeList1.Items;
            rNode := nodes.AddChild(nil, lUIA[0]);
            for i := 0 to 54 do
            begin

                if i < 31 then
                begin
                    if (flgUIA and TruncPow(2, i)) <> 0 then
                    begin
                        if i >= 28 then
                        begin
                            node := nodes.AddChild(rNode, lUIA[i+1]);
                            UIArray := nil;
                            iRes := E_FAIL;
                            if i = 28 then
                            begin
                                iRes := UIEle.Get_CurrentControllerFor(UIArray);
                            end
                            else if i = 29 then
                            begin
                                iRes := UIEle.Get_CurrentDescribedBy(UIArray);
                            end
                            else if i = 30 then
                            begin
                                iRes := UIEle.Get_CurrentFlowsTo(UIArray);
                            end;
                            if Assigned(UIArray) and SUCCEEDED(ires) then
                            begin
                                //caption := '1';
                                if SUCCEEDED(UIArray.Get_Length(iLen)) then
                                begin
                                    //caption := '2-a / ' + inttostr(ilen);
                                    for t := 0 to iLen - 1 do
                                    begin
                                        cNode := nodes.AddChild(Node, inttostr(t));
                                        //caption := '2-b';
                                        ResEle := nil;
                                        if SUCCEEDED(UIArray.GetElement(t, ResEle)) then
                                        begin
                                            //caption := '3';
                                            if Assigned(ResEle) then
                                            begin
                                                //caption := '4';
                                                ws := none;
                                                iaTarg := GetMSAA;
                                                if Assigned(iaTarg) then
                                                begin
                                                    iaTarg.Get_accName(0, ws);
                                                    New(TD);
                                                    TD^.Acc := iaTarg;
                                                    TD^.iID := 0;
                                                    cNode.Data := Pointer(TD);
                                                end;

                                                TreeList1.SetNodeColumn(cnode, 1, ws);
                                                Result := Result + lUIA[i+1] + ':' + #9 + ws + #13#10;
                                            end;
                                        end;
                                    end;
                                end;
                            end
                            else
                            begin
                                TreeList1.SetNodeColumn(node, 1, none);
                            end;
                        end
                        else
                        begin
                            node := nodes.AddChild(rNode, lUIA[i+1]);
                            TreeList1.SetNodeColumn(node, 1, UIAs[i]);
                            Result := Result + lUIA[i+1] + ':' + #9 + UIAs[i] + #13#10;
                        end;
                    end;
                end
                else
                begin
                  if (flgUIA2 and TruncPow(2, i-31)) <> 0 then
                  begin
                      if i = 31 then
                      begin
                        node := nodes.AddChild(rNode, lUIA[i+1]);
                        try
                            iRes := UIEle.Get_CurrentLabeledBy(ResEle);
                            if SUCCEEDED(iRes) then
                            begin


                                if not SUCCEEDED(ResEle.Get_CurrentName(ws)) then
                                    ws := none;
                                //node := nodes.AddChild(node, ws);
                                TreeList1.SetNodeColumn(node, 1, ws);
                                Result := Result + lUIA[i+1] + ':' + #9 + ws + #13#10;
                                 New(TD);
                                 TD^.Acc := GetMSAA;
                                 TD^.iID := 0;
                                 node.Data := Pointer(TD);

                                //UIAs[31] := ws;
                            end
                            else
                                TreeList1.SetNodeColumn(node, 1, none);
                        except
                        on E: Exception do
                            //UIAs[31] := E.Message;
                        end;
                      end
                      else if i = 53 then
                      begin
                      	node := nodes.AddChild(rNode, lUIA[54]);
                        Result := Result + lUIA[54];
                        try
                        	for t := 0 to 5 do
                          begin
                            cNode := nodes.AddChild(Node, lUIA[55+t]);
                          	TreeList1.SetNodeColumn(cnode, 1, UIAs[53+t]);
                            Result := Result + #9 + lUIA[55+t] + ':' + UIAs[53+t] + #13#10;
                          end;
                        except
                        on E: Exception do
                            //UIAs[31] := E.Message;
                        end;
                      end
                      else// if i = 32 then
                      begin

                        node := nodes.AddChild(rNode, lUIA[i+1]);
                        TreeList1.SetNodeColumn(node, 1, UIAs[i]);
                        Result := Result + lUIA[i+1] + ':' + #9 + UIAs[i] + #13#10;
                      end;
                  end;
                end;
            end;
            rNode.Expand(True);
        end;
    finally
    end;

end;

function TwndMSAAV.ARIAText4FF: string;
var
    ARIAs: array [0..1, 0..1] of string;
    s, Path, sARIA: string;
    List, List2: TStringList;
    i: integer;
    aPC, aPC2: array [0..64] of pWidechar;
    aSI: array [0..64] of Smallint;
    WD: Word;
    ini: TMemInifile;
    aAttrs: string;
    rNode, Node: TTreeNode;
begin
    if SDOM = nil then
    begin
        GetNaviState(True);
        Exit;
    end;
    sAria := 'ARIA';
    ARIAs[0, 0] := 'Role';
    ARIAs[1, 0] := 'Attributes';
    if FileExists(Transpath) then
    begin
        ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	    try
            sAria := ini.ReadString('ARIA', 'ARIA', sAria);
            ARIAs[0, 0] := ini.ReadString('ARIA', 'Role', ARIAs[0, 0]);
            ARIAs[1, 0] := ini.ReadString('ARIA', 'Attributes', ARIAs[1, 0]);
        finally
            ini.Free;
        end;
    end;
    try
        Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
        SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
        for i := 0 to WD - 1 do
        begin
            s := LowerCase(aPC[i]);
            if s = 'role' then
            begin
                Arias[0, 1] := aPC2[i];
            end
            else if Copy(s, 1, 4) = 'aria' then
            begin
                aattrs := aattrs + '"' +  WideString(aPC[i]) + '","' + WideString(aPC2[i]) + '"' +  #13#10;
            end;
        end;
        if ARIAs[0, 1] = '' then
            ARIAs[0, 1] := none;
        if aattrs = '' then
            ARIAS[1, 1] := none
        else
            ARIAs[1, 1] := aattrs;
        rNode := TreeList1.Items.AddChild(nil, sAria);
        Node := TreeList1.Items.AddChild(rNode, Arias[0, 0]);
        TreeList1.SetNodeColumn(node, 1, Arias[0, 1]);
        Node := TreeList1.Items.AddChild(rNode, Arias[1, 0]);
        if aattrs = '' then
            TreeList1.SetNodeColumn(node, 1, Arias[1, 1])
        else
        begin
            rNode := Node;
            List := TStringList.Create;
            List2 := TStringList.Create;
            try
                List.Text := aattrs;
                for i := 0 to List.Count - 1 do
                begin
                    List2.Clear;
                    List2.CommaText := List[i];
                    if List2.Count >= 1 then
                    begin
                        Node := TreeList1.Items.AddChild(rNode, List2[0]);
                        if List2.Count >= 2 then
                        begin
                            TreeList1.SetNodeColumn(node, 1, List2[1]);
                        end;
                    end;
                end;
            finally
                List.Free;
                List2.Free;
            end;
        end;
        Result := sAria + #13#10#13#10 + Arias[0, 0] + ':' + #9 +  Arias[0, 1] +
                  #13#10 + Arias[1, 0] + ':' + #9 +  Arias[1, 1] + #13#10#13#10;
        rNode.Expand(True);
    except

    end;
end;

function TwndMSAAV.GetSource(iDoc: IHTMLDOCUMENT): string;
var
    AStream:TMemoryStream;
    List: TStringList;
begin
    AStream:=TMemoryStream.Create;
    try
        (iDoc as IPersistStreamInit).Save(TStreamAdapter.Create(AStream),false);
        List := TStringList.Create;
        try
            AStream.Seek(0,soFromBeginning);
            List.LoadFromStream(AStream);
            Result := List.Text;
            //Result := LoadMemoFromMemoryStream(AStream);
        finally
            List.Free;
        end;
    finally
        AStream.Free;
    end;
end;

procedure TwndMSAAV.SetBalloonPos(X, Y:integer);
begin
    SendMessage(hWndTip, TTM_TRACKPOSITION,0, MAKELONG(X, Y));
    sendMessage(hWndTip, TTM_TRACKACTIVATE, 1, Integer(@ti));
end;

procedure TwndMSAAV.ShowBalloonTip(Control: TWinControl; Icon: integer; Title: string; Text: string; RC: TRect; X, Y:integer; Track:boolean = false);
{const
  TOOLTIPS_CLASS = 'tooltips_class32';
  TTS_ALWAYSTIP = $01;
  TTS_NOPREFIX = $02;
  TTS_BALLOON = $40;
  TTF_SUBCLASS = $0010;
  TTF_TRANSPARENT = $0100;
  TTF_CENTERTIP = $0002;
  TTM_ADDTOOL = $0400 + 50;
  TTM_SETTITLE = (WM_USER + 32);
  ICC_WIN95_CLASSES = $000000FF;
type
  TOOLINFO = packed record
    cbSize: Integer;
    uFlags: Integer;
    hwnd: THandle;
    uId: Integer;
    rect: TRect;
    hinst: THandle;
    lpszText: PWideChar;
    lParam: Integer;
  end;  }
var
  //hWndTip: THandle;

  hhWnd: THandle;
  TP1: TPoint;
  Wnd: hwnd;
begin
    WindowFromAccessibleObject(iAcc, Wnd);
  hhWnd    := Control.Handle;
  hWndTip := CreateWindow(TOOLTIPS_CLASS, nil,
    WS_POPUP or TTS_NOPREFIX or TTS_BALLOON or TTS_ALWAYSTIP ,//or TTS_CLOSE ,
    RC.Left, RC.Top, 0, 0, hhWnd, 0, HInstance, nil);
  if hWndTip <> 0 then
  begin
    {SetWindowPos(hWndTip, HWND_TOPMOST, 0, 0, 0, 0,
      SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);  }

    ti.cbSize := SizeOf(ti);
    if not Track then
        ti.uFlags := {TTF_CENTERTIP or }{TTF_IDISHWND  or }TTF_TRANSPARENT or TTF_SUBCLASS {or TTF_TRACK or TTF_ABSOLUTE}
    else
        ti.uFlags := {TTF_CENTERTIP or }{TTF_IDISHWND  or }TTF_TRANSPARENT or TTF_SUBCLASS or TTF_TRACK or TTF_ABSOLUTE;
    ti.hwnd := wnd{hhWnd};
    ti.uId := 1;
    ti.lpszText := PChar(Text);
    //Windows.GetClientRect(hhWnd, ti.rect);
    ti.Rect := RC;
    //SendMessage(hWndTip, TTM_SETTIPBKCOLOR, BackCL, 0);
    //SendMessage(hWndTip, TTM_SETTIPTEXTCOLOR, TextCL, 0);

    SendMessageW(hWndTip,TTM_SETMAXTIPWIDTH,0,640);
    SendMessage(hWndTip, TTM_ADDTOOL, 1, Integer(@ti));
    SendMessage(hWndTip, TTM_SETTITLE, Icon,  Integer(PChar(title)));
    //SendMessageW(hWndTip,TTM_UPDATETIPTEXTW,0,Integer(@ti));

    TP1.X := Control.Left;
    TP1.Y := Control.Top;
    TP1 := Control.ClientToScreen(TP1);
    if not Track then
    begin
        SendMessage(hWndTip, TTM_TRACKPOSITION,0, MAKELONG(X, Y));
        sendMessage(hWndTip, TTM_TRACKACTIVATE, 1, Integer(@ti));
    end;
  end;
end;

procedure ShowException(AHandle: HWND; ATitle, AText: string; Icon: Integer);

var

  BaloonTip: TEditBalloonTip;

begin

  BaloonTip.cbStruct := SizeOf(TEditBalloonTip);

  BaloonTip.pszTitle := PChar(ATitle);

  BaloonTip.pszText := PChar(AText);

  BaloonTip.ttiIcon := TTI_INFO_LARGE;//Icon;


  SendMessage(AHandle, EM_SHOWBALLOONTIP, 0, Integer(@BaloonTip));

end;

procedure TwndMSAAV.acShowTipExecute(Sender: TObject);
begin
    acShowTip.Checked := not acShowTip.Checked;
    if acShowTip.Checked then
    begin
        {if not Assigned(WndTip) then
            WndTip := TfrmTipWnd.Create(self); }
        ShowTipWnd;
    end
    else
    begin
        if WndTip <> nil then
        begin
            SendMessage(WndTip.TipInfo.Handle, EM_HIDEBALLOONTIP, 0, 0);
            FreeAndNil(WndTip);
            WndTip := nil;
        end;
        if hWndTip <> 0 then
            DestroyWindow(hWndTip);

    end;
end;

procedure TwndMSAAV.ShowTipWnd;
var
    RC: TRect;
    vChild: variant;
    Mon : TMonitor;
    s, src, c, inner, outer, sTxt: string;
    sLeft, sTop, i, iRes: integer;
    iSP: IServiceProvider;
    iEle: IHTMLElement;
    isd: ISimpleDOMNode;
    PC: pchar;
    ini: TMemInifile;
    PC2:PChar;
    SI: Smallint;
    PU, PU2: PUINT;
    WD: Word;
    iText: ISimpleDOMText;
    tinfo: TToolInfo;
    wnd: HWND;
begin
    if not Assigned(iAcc) then Exit;
    try



        if mnublnMSAA.Checked then
            s := s + MSAAText4Tip;
        if mnublnIA2.Checked then
            s := s + SetIA2Text(iacc, false);
        if mnublnCode.Checked then
        begin
            ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
	        try

            c := ini.ReadString('HTML', 'Code', 'Code');
            inner := '(' + ini.ReadString('HTML', 'innerHTML', 'innerHTML') + ')';
            outer := '(' + ini.ReadString('HTML', 'outerHTML', 'outerHTML') + ')';
            sTxt := ini.ReadString('HTML', 'Text', 'Text');
            finally
                ini.Free;
            end;
            iSP := iAcc as IServiceProvider;
            if iMode = 0 then
            begin


                if (SUCCEEDED(iSP.QueryService(IID_IHTMLELEMENT, IID_IHTMLELEMENT, iEle))) then
                begin
                    Src := iEle.outerHTML;
                    Src := Copy(Src, 0, ShowSrcLen);
                    s := s + c + outer + ':' + #13#10 + Src;
                end;
            end
            else if iMode = 1 then
            begin
                if SUCCEEDED(iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isd)) then
                begin
                    isd.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);

                    if WD <>3 then
                    begin
                        PC := '';
                        isd.get_innerHTML(PC);
                        Src := Copy(PC, 0, ShowSrcLen);
                        s := s + c +inner  + #13#10 + Src;
                    end
                    else
                    begin
                        iSP := isd as IServiceProvider;
                        if SUCCEEDED(iSP.QueryInterface(IID_ISIMPLEDOMTEXT, iText)) then
                        begin
                            iText.get_domText(PC);
                            Src := Copy(PC, 0, ShowSrcLen);
                            s := s + c +'(' + sTxt + ')' + #13#10 + Src;
                        end;
                    end;
                end;
            end;
            //s := s + Src;
        end;


        vChild := varParent;//CHILDID_SELF;
        iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);

        sLeft := RC.Left;
        sTop := RC.Top;// + RC.Bottom;

        if hWndTip <> 0 then
        begin
            DestroyWindow(hWndTip);
        end;
            ShowBalloonTip(self, 1, 'Aviewer', s, rc, sLeft, sTop, True);
            SetBalloonPos(sLeft, sTop);
            WindowFromAccessibleObject(iAcc, Wnd);
            tinfo.cbSize := SizeOf(tinfo);
            tinfo.hwnd := Wnd;
            tinfo.uId := 1;
            iRes := SendMessage(hWndTip, TTM_GETBUBBLESIZE , 0, Integer(@tinfo));

            if iRes > 0 then
            begin
                i := HIWORD(iRes);
                Mon := Screen.MonitorFromRect(rc);

                    if ((RC.Top + i + 20) > {Mon.WorkareaRect.Top + }Mon.WorkareaRect.Bottom) then
                        sTop := RC.Top - i - 20// - RC2.top)
                    else if (RC.Top + RC.Bottom) > mon.WorkareaRect.Bottom then
                        sTop := RC.Top - i - 20
                    else
                        sTop := RC.Top + RC.Bottom;
                    if sTop < 0 then
                      sTop := 0;
                       OutputDebugString(PWIDECHAR(inttostr(rc.Bottom)));
                       OutputDebugString(PWidechar(inttostr(RC.Top + i + 20)));
                      if (RC.Left + LOWORD(iRes)) > (Mon.WorkareaRect.Left + Mon.WorkareaRect.Right) then
                          sLeft := Mon.WorkareaRect.Right - LOWORD(iRes)
                      else if RC.Left < Mon.WorkareaRect.Left then
                          sLeft := Mon.WorkareaRect.Left
                      else
                          sLeft := RC.Left;
                     //caption := 'T:' + inttostr(stop) + '/L:' + inttostr(sleft);
                    //ShowBalloonTip(self{WndTip.TipInfo}, 1, 'Aviewer', s, RC, sLeft, sTop);
                    SetBalloonPos(sLeft, sTop);
            end
            else
            begin
                SetBalloonPos(sLeft, sTop);
            end;

    except
        on E:Exception do
        begin
            ShowErr(E.Message);
        end;
    end;

end;

procedure TwndMSAAV.ShowRectWnd(bkClr: TColor);
var
    RC: TRect;
    vChild: variant;
begin
    if not Assigned(iAcc) then Exit;

    try
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        WndFocus.Shape1.Brush.Color := bkClr;
        vChild := varParent;//CHILDID_SELF;
        iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
        SetWindowPos(WndFocus.Handle, HWND_TOPMOST, RC.Left - 5, RC.Top - 5, RC.Right + 10, RC.Bottom + 10, SWP_SHOWWINDOW or SWP_NOACTIVATE);
        if hRgn1 <> 0 then
        begin
            DeleteObject(hRgn1);
            DeleteObject(hRgn2);
            DeleteObject(hRgn3);
            hRgn1 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
            SetWindowRgn(WndFocus.Handle, hRgn1, False);
            InvalidateRect(WndFocus.Handle, nil, false);
        end;
        hRgn1 := CreateRectRgn(0, 0, 1, 1);
        hRgn2 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
        hRgn3 := CreateRectRgn(5, 5, WndFocus.Width - 5, WndFocus.Height - 5);
        CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
        SetWindowRgn(WndFocus.Handle, hRgn1, TRUE);
        WndFocus.Visible := true;
    except

    end;
end;

procedure TwndMSAAV.ShowRectWnd2(bkClr: TColor; RC:TRect);
var
    vChild: variant;
begin
    if not Assigned(iAcc) then Exit;

    try
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        WndFocus.Shape1.Brush.Color := bkClr;
        vChild := varParent;//CHILDID_SELF;
        //iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
        SetWindowPos(WndFocus.Handle, HWND_TOPMOST, RC.Left - 5, RC.Top - 5, RC.Right + 10, RC.Bottom + 10, SWP_SHOWWINDOW or SWP_NOACTIVATE);
        if hRgn1 <> 0 then
        begin
            DeleteObject(hRgn1);
            DeleteObject(hRgn2);
            DeleteObject(hRgn3);
            hRgn1 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
            SetWindowRgn(WndFocus.Handle, hRgn1, False);
            InvalidateRect(WndFocus.Handle, nil, false);
        end;
        hRgn1 := CreateRectRgn(0, 0, 1, 1);
        hRgn2 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
        hRgn3 := CreateRectRgn(5, 5, WndFocus.Width - 5, WndFocus.Height - 5);
        CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
        SetWindowRgn(WndFocus.Handle, hRgn1, TRUE);
        WndFocus.Visible := true;
    except

    end;
end;

procedure TwndMSAAV.ShowDescWnd(bkClr: TColor; RC:TRect);
var
    vChild: variant;
begin
    if not Assigned(iAcc) then Exit;

    try
        if not Assigned(WndDesc) then
            WndDesc := TWndFocusRect.Create(self);
        WndDesc.Shape1.Brush.Color := bkClr;
        vChild := varParent;//CHILDID_SELF;
        //iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
        SetWindowPos(WndDesc.Handle, HWND_TOPMOST, RC.Left - 10, RC.Top - 10, RC.Right + 20, RC.Bottom + 20, SWP_SHOWWINDOW or SWP_NOACTIVATE);
        if hRgn1 <> 0 then
        begin
            DeleteObject(hRgn1);
            DeleteObject(hRgn2);
            DeleteObject(hRgn3);
            hRgn1 := CreateRectRgn(0, 0, WndDesc.Width, WndDesc.Height);
            SetWindowRgn(WndDesc.Handle, hRgn1, False);
            InvalidateRect(WndDesc.Handle, nil, false);
        end;
        hRgn1 := CreateRectRgn(0, 0, 1, 1);
        hRgn2 := CreateRectRgn(0, 0, WndDesc.Width, WndDesc.Height);
        hRgn3 := CreateRectRgn(5, 5, WndDesc.Width - 5, WndDesc.Height - 5);
        CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
        SetWindowRgn(WndDesc.Handle, hRgn1, TRUE);
        WndDesc.Visible := true;
    except

    end;
end;
procedure TwndMSAAV.ShowLabeledWnd(bkClr: TColor; RC:TRect);
var
    vChild: variant;
begin
    if not Assigned(iAcc) then Exit;

    try
        if not Assigned(WndLabel) then
            WndLabel := TWndFocusRect.Create(self);
        WndLabel.Shape1.Brush.Color := bkClr;
        vChild := varParent;//CHILDID_SELF;
        //iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
        SetWindowPos(WndLabel.Handle, HWND_TOPMOST, RC.Left - 10, RC.Top - 10, RC.Right + 20, RC.Bottom + 20, SWP_SHOWWINDOW or SWP_NOACTIVATE);
        if hRgn1 <> 0 then
        begin
            DeleteObject(hRgn1);
            DeleteObject(hRgn2);
            DeleteObject(hRgn3);
            hRgn1 := CreateRectRgn(0, 0, WndLabel.Width, WndLabel.Height);
            SetWindowRgn(WndLabel.Handle, hRgn1, False);
            InvalidateRect(WndLabel.Handle, nil, false);
        end;
        hRgn1 := CreateRectRgn(0, 0, 1, 1);
        hRgn2 := CreateRectRgn(0, 0, WndLabel.Width, WndLabel.Height);
        hRgn3 := CreateRectRgn(5, 5, WndLabel.Width - 5, WndLabel.Height - 5);
        CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
        SetWindowRgn(WndLabel.Handle, hRgn1, TRUE);
        WndLabel.Visible := true;
    except

    end;
end;

procedure TwndMSAAV.ShowTargWnd(bkClr: TColor; RC:TRect);
var
    vChild: variant;
begin
    if not Assigned(iAcc) then Exit;

    try
        if not Assigned(WndTarg) then
            WndTarg := TWndFocusRect.Create(self);
        WndTarg.Shape1.Brush.Color := bkClr;
        vChild := varParent;//CHILDID_SELF;
        //iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
        SetWindowPos(WndTarg.Handle, HWND_TOPMOST, RC.Left - 10, RC.Top - 10, RC.Right + 20, RC.Bottom + 20, SWP_SHOWWINDOW or SWP_NOACTIVATE);
        if hRgn1 <> 0 then
        begin
            DeleteObject(hRgn1);
            DeleteObject(hRgn2);
            DeleteObject(hRgn3);
            hRgn1 := CreateRectRgn(0, 0, WndTarg.Width, WndTarg.Height);
            SetWindowRgn(WndTarg.Handle, hRgn1, False);
            InvalidateRect(WndTarg.Handle, nil, false);
        end;
        hRgn1 := CreateRectRgn(0, 0, 1, 1);
        hRgn2 := CreateRectRgn(0, 0, WndTarg.Width, WndTarg.Height);
        hRgn3 := CreateRectRgn(5, 5, WndTarg.Width - 5, WndTarg.Height - 5);
        CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
        SetWindowRgn(WndTarg.Handle, hRgn1, TRUE);
        WndTarg.Visible := true;
    except

    end;
end;

function TwndMSAAV.GetWindowNameLC(Wnd: HWND): string;
var
    Len:integer;
    PC:PChar;
    s: string;
begin
    GetMem(PC,100);
    Len := GetClassName(Wnd,PC,100);
    SetString(s,PC,Len);
    FreeMem(PC);
    result := LowerCase(s);
end;

function TwndMSAAV.GetOpeMode(CName: string): integer;
var
  i: integer;
begin
  Result := 2;
  if (LowerCase(CName) = 'internet explorer_server') or (LowerCase(CName) = 'macromediaflashplayeractivex') then
    Result := 0
  else
  begin
    for i := Low(CNames) to High(CNames) do
    begin
      if (CNames[i] = LowerCase(CName)) then
      begin
        Result := 1;
        break;
      end;
    end;
  end;
  {else if (s = 'mozillauiwindowclass') or (s = 'chrome_renderwidgethosthwnd') or (s = 'mozillawindowclass') or (s = 'chrome_widgetwin_0') or (s = 'chrome_widgetwin_1') then
    Result := 1
  else
    Result := 2; }
end;



procedure TwndMSAAV.Timer1Timer(Sender: TObject);
var
    Wnd: hwnd;
    i: cardinal;
    pAcc: IAccessible;
    v: variant;
    idis: iDispatch;
    po: tagPOINT;
    //PT:TPoint;
begin
    if not acCursor.Checked then
        exit;
    arPT[2] := arPT[1];
    arPT[1] := arPT[0];
    getcursorpos(arPT[0]);

    if (arPT[0].x = oldPT.X) and (arPT[0].Y = oldPT.Y) then
        Exit;
    try
    if (arPT[0].x = arPT[1].x) and (arPT[0].x = arPT[2].x) and (arPT[2].x = arPT[1].x) and
    (arPT[0].y = arPT[1].y) and (arPT[0].y = arPT[2].y) and (arPT[2].y = arPT[1].y) then
    begin
        if SUCCEEDED(AccessibleObjectFrompoint(arPT[0], @pAcc, v)) then
        begin
            if ((not acOnlyFocus.Checked) and (acCursor.Checked)) then
            begin
                if SUCCEEDED(WindowFromAccessibleObject(pAcc, Wnd)) then
                begin
                    po.x := arPT[0].X;
                    po.y := arPT[0].y;
                    UIAuto.ElementFromPoint(po, UIEle);
                    i := GetWindowLong(Wnd, GWL_HINSTANCE);
                    if i = hInstance then
                    begin
                        pAcc := nil;
                        exit;
                    end;
                    //if IsSameUIElement(iAcc, pAcc) then
                        //Exit;
                    if ( not acOnlyFOcus.Checked) then
                    begin
                        //UIEle := nil;
                        //GetPhysicalCursorPos(PT);
                        //UIA.ElementFromPoint(tagPoint(PT), UIEle);
                        //UIA.ElementFromIAccessible(pAcc, 0, UIEle);
                        Treemode := false;
                        iAcc := pAcc;
                        //iDis := pAcc.accParent;
                        pAcc.Get_accParent(iDis);
                        accRoot := iDis as IAccessible;
                        varparent := v;
                        if (acCursor.Checked) then
                        begin
                            //s := GetWindowNameLC(Wnd);
                            iMode := GetOpeMode(GetWindowNameLC(Wnd));

                            //showmessage(s);
                            //Memo1.Lines.Add(s);
                            if (not acMSAAMode.Checked) then
                            begin
                                if iMode <> 2 then
                                begin
                                    if (acRect.Checked) then
                                    begin
                                        ShowRectWnd(clRed);
                                    end;

                                    ShowMSAAText;
                                    {if (acShowTip.Checked) then
                                    begin
                                        ShowTipWnd;
                                    end; }
                                end
                                else
                                    iAcc := nil;
                            end
                            else
                            begin
                                if (acRect.Checked) then
                                begin
                                    ShowRectWnd(clRed);
                                end;

                                ShowMSAAText;
                                {if (acShowTip.Checked) then
                                begin
                                    ShowTipWnd;
                                end;}
                            end;
                            if iAcc <> nil then
                            begin
                                OldPT := arPT[0];
                                arPT[0] := Point(0, 0);
                                arPT[1] := Point(0, 0);
                                arPT[2] := Point(0, 0);
                            end;
                        end;
                    end
                    else
                        iAcc := nil;

                end;
            end
            else
                iAcc := nil;

        end;
    end;
    except

    end;

end;


procedure TwndMSAAV.Timer2Timer(Sender: TObject);
begin
    Timer2.Enabled := false;
    Treemode := false;
    if (not wndMSAAV.acMSAAMode.Checked) then
    begin
        if ( not wndMSAAV.acOnlyFOcus.Checked) then
        begin


            wndMSAAV.TreeList1.Items.BeginUpdate;
            wndMSAAV.TreeList1.Items.Clear;
            wndMSAAV.ShowMSAAText;

            wndMSAAV.TreeList1.Items.EndUpdate;
        end;
    end
    else
    begin
        if (wndMSAAV.acFocus.Checked) then
        begin

            wndMSAAV.TreeList1.Items.BeginUpdate;
            wndMSAAV.TreeList1.Items.Clear;
            wndMSAAV.ShowMSAAText;

            wndMSAAV.TreeList1.Items.EndUpdate;
        end;
    end;
end;



procedure TwndMSAAV.FormClose(Sender: TObject; var Action: TCloseAction);
var
    ini: TMemINiFile;
    i: integer;
begin
    //UIAuto.Free;
    UIAuto := nil;
    CoUnInitialize;
    NotifyWinEvent(EVENT_OBJECT_DESTROY, Handle, OBJID_CLIENT, 0);

    if hHook <> 0 then
                begin
                    if UnhookWinEvent (hHook) then
                    begin
                        //FreeHookInstance(HookProc);
                        hHook := 0;
                    end;
                end;
    try
    ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
    try
        ini.WriteInteger('Settings', 'Width', width);
        ini.WriteInteger('Settings', 'Height', Height);
        ini.WriteInteger('Settings', 'Top', Top);
        ini.WriteInteger('Settings', 'Left', Left);
        //P2W, P4H: integer;
        ini.WriteInteger('Settings', 'P2W', Panel2.Width);
        ini.WriteInteger('Settings', 'P4H', Panel4.Height);
        ini.WriteBool('Settings', 'vMSAA', mnuMSAA.Checked);
        ini.WriteBool('Settings', 'vARIA', mnuARIA.Checked);
        ini.WriteBool('Settings', 'vHTML', mnuHTML.Checked);
        ini.WriteBool('Settings', 'vIA2', mnuIA2.Checked);
        ini.WriteBool('Settings', 'vUIA', mnuUIA.Checked);
        ini.WriteBool('Settings', 'TVAll', mnuAll.Checked);

        ini.WriteBool('Settings', 'bMSAA', mnublnMSAA.Checked);
        ini.WriteBool('Settings', 'bIA2', mnublnIA2.Checked);
        ini.WriteBool('Settings', 'bCode', mnublnCode.Checked);
        ini.UpdateFile;
    finally
        ini.Free;
    end;
    if CP <> nil then
        CP.Unadvise(Cookie);
    for i := 0 to FCPList.Count - 1 do
    begin
        if Assigned(FCPList.Items[i]) then
            Dispose(PCP(FCPList.Items[i]));
    end;
    for i := 0 to FIEList.Count - 1 do
    begin
        if Assigned(FIEList.Items[i]) then
            Dispose(PIE(FIEList.Items[i]));
    end;
    for i := 0 to FCURList.Count - 1 do
    begin
        if Assigned(FCURList.Items[i]) then
            Dispose(PUIA(FCURList.Items[i]));
    end;
    for i := 0 to FCACList.Count - 1 do
    begin
        if Assigned(FCACList.Items[i]) then
            Dispose(PUIA(FCACList.Items[i]));
    end;

    FCPList.Clear;
    FCPList.Free;
    FIEList.Clear;
    FIEList.Free;
    FCURList.Clear;
    FCURList.Free;
    FCACList.Clear;
    FCACList.Free;
    TBList.Free;
    LangList.Free;
    if hWndTip <> 0 then
        DestroyWindow(hWndTip);
    except

    end;
end;





procedure TwndMSAAV.FormCreate(Sender: TObject);
var
    Rec     : TSearchRec;
    i: integer;

begin



    TreeTH := nil;
    APPDir :=  IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
    TransDir := IncludeTrailingPathDelimiter(AppDir + 'Languages');
    if not DirectoryExists(TransDir) then
      TransDir := IncludeTrailingPathDelimiter(AppDir + 'Lang');

    mnuLang.Visible := FileExists(TransDir + 'Default.ini');
    if mnuLang.Visible then
        Transpath := TransDir + 'Default.ini'
    else
        Transpath := APPDir + ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
    DllPath := APPDir + 'IAccessible2Proxy.dll';
    //memo1.ParentWindow := handle;
    arPT[0] := Point(0, 0);
    arPT[1] := Point(0, 0);
    arPT[2] := Point(0, 0);
    DMode := False;
    for i := 1 to ParamCount do
    begin
        if LowerCase(ParamStr(i)) = '-d' then
            DMode := True;

    end;
    //DoubleBuffered := True;
    FCPList := TList.Create;
    FIEList := TList.Create;
    FCURList := TList.Create;
    FCACList := TList.Create;
    TBList := TIntegerList.Create;
    LangList := TStringList.Create;
    Created := True;
    SPath := IncludeTrailingPathDelimiter(GetMyDocPath) + 'MSAAV.ini';
    tbParent.Enabled := false;
    tbChild.Enabled := false;
    tbPrevS.Enabled := false;
    tbNextS.Enabled := false;
    acParent.Enabled := tbParent.Enabled;
    acChild.Enabled := tbChild.Enabled;
    acPrevS.Enabled := tbPrevS.Enabled;
    acNextS.Enabled := tbNextS.Enabled;
    iFocus := 1;
    StopWait := False;
    if mnuLang.Visible then
    begin
        if  (FindFirst(TransDir + '*.ini', faAnyFile, Rec) = 0) then
        begin
            repeat
                if  ((Rec.Name <> '.') and (Rec.Name <> '..')) then
                begin
                    if  ((Rec.Attr and faDirectory) = 0)  then
                    begin
                    LangList.Add(Rec.Name);
                    end;
                end;
            until (FindNext(Rec) <> 0);
        end;
    end;

    Load;
    ExecCmdLine;
end;


procedure TwndMSAAV.FormDestroy(Sender: TObject);
begin
    if Assigned(TreeTH) then TreeTH.Free;

end;

procedure TwndMSAAV.FormPaint(Sender: TObject);
begin
    {if Created then
    begin
        Created := False;
        CopyDataToOld(Handle);

    end;  }

end;

procedure TwndMSAAV.FormResize(Sender: TObject);
begin

    P2W := Panel2.Width;
    P4H := Panel4.Height;
end;


procedure TwndMSAAV.TreeList1Change(Sender: TObject; Node: TTreeNode);
var
    RC: TRect;
    iSP: iServiceProvider;
    Ele: IHTMLElement;
    sDom: ISimpleDOMNode;
begin

    if TreeList1.Selected <> nil then
    begin

        if Node.Data <> nil then
        begin
            if acRect.Checked then
            begin
                if iMode = 0 then
                begin
                    iSP := iAcc as IServiceProvider;
                    if (SUCCEEDED(iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, Ele))) then
                        Ele.scrollIntoView(True);
                end
                else if iMode = 1 then
                begin
                    iSP := iAcc as IServiceProvider;
                    if (SUCCEEDED(iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, sDom))) then
                        sDom.scrollTo(True);
                end;
                Sleep(50);
                TTreeData(Node.Data^).Acc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0);
                ShowTargWnd(clBlue, RC);
            end;
        end;


    end;
end;

function TwndMSAAV.AccIsNull(tAcc: IAccessible): boolean;
var
    UIEle1: IUIAutomationElement;
begin
    result := True;
    if Assigned(tAcc) then
    begin
        if Assigned(UIAuto) then
        begin
            if SUCCEEDED(UIAuto.ElementFromIAccessible(tAcc, varParent, UIEle1)) then
                Result := False;
        end;
    end;
end;
procedure TwndMSAAV.TreeList1Click(Sender: TObject);
begin
    if TreeList1.Selected <> nil then
    begin
        if Assigned(UIAuto) and Assigned(iAcc)then
        begin
            if not AccIsNull(iAcc) then
            begin
                if iDefIndex = TreeList1.Selected.Index then
                begin
                    try
                        iAcc.accDoDefaultAction(VarParent);
                    except

                    end;
                end;
            end
            else
                iAcc := nil;
        end;
    end;
end;

procedure TwndMSAAV.TreeList1Deletion(Sender: TObject; Node: TTreeNode);
begin
    Dispose(Node.Data);
end;

procedure TwndMSAAV.TreeList1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (TreeList1.Selected <> nil) and (key = VK_RETURN) then
    begin
        if Assigned(UIAuto) and Assigned(iAcc)then
        begin
            if not AccIsNull(iAcc) then
            begin
                if iDefIndex = TreeList1.Selected.Index then
                begin
                    try
                        iAcc.accDoDefaultAction(VarParent);
                    except

                    end;
                end;
            end
            else
                iAcc := nil;
        end;
    end;
end;

//procedure TwndMSAAV.TreeView1Addition(Sender: TObject; Node: TTreeNode);
//var
//    ovChild: Olevariant;
//    function Get_RoleText(Acc: IAccessible; Child: integer): string;
//    var
//        PC:PChar;
//        ovValue: OleVariant;
//    begin
//        ovChild := Child;
//        //ovValue := Acc.accRole[ovChild];
//        try
//                    //if (IDispatch(ovValue) = nil) then
//            if not VarHaveValue(ovValue) then
//            begin
//                Result := None;
//            end
//            else
//            begin
//                if VarIsNumeric(ovValue) then
//                begin
//                    PC := StrAlloc(255);
//                    GetRoleTextW(ovValue, PC, StrBufSize(PC));
//                    Result := PC;
//                    StrDispose(PC);
//                end
//                else if VarIsStr(ovValue) then
//                begin
//                    Result := VarToStr(ovValue);
//                end;
//            end;
//        except
//            Result := None;
//        end;
//    end;
//begin
//
//    try
//        if (node.Data <> nil) and (TTreeData(Node.Data^).Acc <> nil) and (Assigned(TreeTH)) then
//        begin
//            Application.ProcessMessages;
//            ovChild := TTreeData(Node.Data^).iID;
//            TTreeData(Node.Data^).Acc.Get_accName(ovChild, ws);
//            if ws = '' then ws := None;
//            Role := Get_ROLETExt(TTreeData(Node.Data^).Acc, ovChild);
//            node.Text := ws + ' - ' + Role;
//            TTreeData(Node.Data^).Acc.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, ovChild);
//            if (sRC.Left = cRC.Left) and (sRC.Top = cRC.Top) and (sRC.Right = cRC.Right) and (sRC.Bottom = cRC.Bottom) then
//            begin
//
//                //dMSAAV.TreeView1.SetFocus;
//                wndMSAAV.TreeView1.TopItem := node;
//
//                node.Selected := True;
//                node.Expanded := true;
//                wndMSAAV.GetNaviState;
//            end;
//        end;
//    except
//
//    end;
//end;

procedure TwndMSAAV.SetTreeMode(pNode: TTreeNode);
var
    b: boolean;
begin
    if TTreeData(pNode.Data^).Acc <> nil then
    begin
        varParent := TTreeData(pNode.Data^).iID;
        if not AccIsNull(TTreeData(pNode.Data^).Acc) then
        begin

                //TreeView1.Enabled := false;
                Treemode := True;
                iAcc := TTreeData(pNode.Data^).Acc;


                if (not acMSAAMode.Checked) then
                begin

                    if iMode <> 2 then
                    begin
                        b := ShowMSAAText;
                        if (acRect.Checked) then
                        begin
                            ShowRectWnd(clRed);
                        end;
                        if b then
                        begin
                            if (acShowTip.Checked) then
                            begin
                                ShowTipWnd;
                            end;
                        end;
                    end
                    else
                        iAcc := nil;
                end
                else
                begin
                    b := ShowMSAAText;
                    if (acRect.Checked) then
                    begin
                        ShowRectWnd(clRed);
                    end;
                    if b then
                    begin
                        if (acShowTip.Checked) then
                        begin
                            ShowTipWnd;
                        end;
                    end;
                end;
            end;

    end;
end;

procedure TwndMSAAV.TreeView1Addition(Sender: TObject; Node: TTreeNode);
var
    Role: string;
    ovChild, ovRole: OleVariant;
    ws: widestring;

begin

 	//tAcc := GetMSAA;
  //TTreeData(Node.Data^).Acc := tAcc;
  //if Node.Text = '' then
  	//if Node.IsVisible then
    //begin
    mnuTVSAll.Enabled := True;
    Node.ImageIndex := 61;
    node.ExpandedImageIndex := 61;
    Node.SelectedIndex := 61;
        TTreeData(Node.Data^).Acc.Get_accName(TTreeData(Node.Data^).iID, ws);
        if ws = '' then ws := None;
        Role := Get_ROLETExt(TTreeData(Node.Data^).Acc, TTreeData(Node.Data^).iID);
        Node.Text := ws + ' - ' + Role;
        if SUCCEEDED(TTreeData(Node.Data^).Acc.Get_accRole(ovChild, ovRole)) then
            begin
                if VarHaveValue(ovRole) then
                begin
                    if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
                    begin
                        Node.ImageIndex := TVarData(ovRole).VInteger - 1;
                        Node.ExpandedImageIndex := Node.ImageIndex ;
                        Node.SelectedIndex := Node.ImageIndex ;
                        //showmessage(inttostr(rNode.ImageIndex));
                    end;
                end;
            end;
    //end;

end;

procedure TwndMSAAV.TreeView1Change(Sender: TObject; Node: TTreeNode);
begin

    if TreeView1.Items.Count <= 0 then
        Exit;
    if TreeView1.SelectionCount = 0 then
    	mnuTVSSel.Enabled := False
    else
    	mnuTVSSel.Enabled := True;
    try
        if (not Assigned(TreeTH)) or (TreeTH.Finished)  then
        begin
            SetTreeMode(Node);

        end;
    finally
        //TreeView1.Enabled := True;
    end;
end;

procedure TwndMSAAV.TreeView1Deletion(Sender: TObject; Node: TTreeNode);
begin
    Dispose(Node.Data);
end;



procedure TwndMSAAV.TreeView1Expanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
var
    i: integer;
    Role: string;
    ovChild, ovRole: OleVariant;
    ws: widestring;

begin

    try
    if Node.Text = '' then
    begin
        TTreeData(Node.Data^).Acc.Get_accName(TTreeData(Node.Data^).iID, ws);
        if ws = '' then ws := None;
        Role := Get_ROLETExt(TTreeData(Node.Data^).Acc, TTreeData(Node.Data^).iID);
        Node.Text := ws + ' - ' + Role;
        if SUCCEEDED(TTreeData(Node.Data^).Acc.Get_accRole(ovChild, ovRole)) then
            begin
                if VarHaveValue(ovRole) then
                begin
                    if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
                    begin
                        Node.ImageIndex := TVarData(ovRole).VInteger - 1;
                        Node.ExpandedImageIndex := Node.ImageIndex ;
                        Node.SelectedIndex := Node.ImageIndex ;
                        //showmessage(inttostr(rNode.ImageIndex));
                    end;
                end;
            end;
    end;
    for i := 0 to node.Count - 1 do
    begin

        if node.Item[i].Text = '' then
        begin
            //if not AccIsNull(TTreeData(Node.Data^).Acc) then
            //begin
            ws := '';
            Role := '';
            //ws := TTreeData(node.Item[i].Data^).Acc.accName[TTreeData(node.Item[i].Data^).iID];
            TTreeData(node.Item[i].Data^).Acc.Get_accName(TTreeData(node.Item[i].Data^).iID, ws);
            if ws = '' then ws := None;
            Role := Get_ROLETExt(TTreeData(node.Item[i].Data^).Acc, TTreeData(node.Item[i].Data^).iID);
            node.Item[i].Text := ws + ' - ' + Role;
            if SUCCEEDED(TTreeData(node.Item[i].Data^).Acc.Get_accRole(ovChild, ovRole)) then
            begin
                if VarHaveValue(ovRole) then
                begin
                    if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
                    begin
                        node.Item[i].ImageIndex := TVarData(ovRole).VInteger - 1;
                        node.Item[i].ExpandedImageIndex := node.Item[i].ImageIndex ;
                        node.Item[i].SelectedIndex := node.Item[i].ImageIndex ;
                        //showmessage(inttostr(rNode.ImageIndex));
                    end;
                end;
            end;
            //end;
        end;
    end;
    finally
        AllowExpansion := True;
    end;
end;

procedure TwndMSAAV.Toolbar1Enter(Sender: TObject);
begin
    Toolbar1.iFocus := 1;
    Toolbar1.Setfocus;
end;

procedure TwndMSAAV.Toolbar1Exit(Sender: TObject);
begin
    Toolbar1.iFocus := 1;
end;



procedure TwndMSAAV.acCopyExecute(Sender: TObject);
var
    Mem: TMemoryStream;
    sList: TStringList;
begin
    //Clipboard.AsText := CPTXT;
    Mem := TMemoryStream.Create;
    sList := TStringList.Create;
    try
        TreeList1.SaveToStream(Mem);
        Mem.Seek(0, soFromBeginning);
        sList.LoadFromStream(Mem);
        Clipboard.AsText := sList.Text;
    finally
        sList.Free;
        Mem.Free;
    end;
end;

procedure TwndMSAAV.acCursorExecute(Sender: TObject);
begin
    acCursor.Checked := not acCursor.Checked;
end;

procedure TwndMSAAV.acFocusExecute(Sender: TObject);
begin
    acFocus.Checked := not acFocus.Checked;
end;



procedure TwndMSAAV.acHelpExecute(Sender: TObject);
begin
    if HelpURL = '' then
        Exit
    else
    begin
        ShellExecuteW(Handle, 'open', PWideChar(HelpURL), nil, nil, SW_SHOW);
    end;
end;



procedure TwndMSAAV.acMSAAModeExecute(Sender: TObject);
begin
    acMSAAMode.Checked := not acMSAAMode.Checked;
    mnuSelD.Enabled := not acMSAAMode.Checked;
    if Assigned(WndFocus) then
    begin
        WndFocus.Visible := false;
    end;
    if Assigned(WndTip) then
    begin
        WndTip.Visible := false;
    end;
    if acMSAAMode.Checked then
    begin
        mnuMSAA.Checked := True;
        mnuARIA.Checked := false;
        mnuHTML.Checked := false;
    end;
end;

procedure TwndMSAAV.ExecFFArrow(sender: TObject);
var
    RC: Trect;
    Serv: IServiceProvider;
    pEle, ppEle: ISimpleDOMNode;
    PC, PC2: PWideChar;
    SI: SmallInt;
    PU, PU2: PUINT;
    WD: Word;
    tc: TComponent;
    s: string;
begin
    if SDOM = nil then
    begin
        GetNaviState(True);
        Exit;
    end;

    if (Sender is TComponent) then
    begin
        tc := Sender as TComponent;
    end
    else
        Exit;
    s := LowerCase(tc.Name);
    if (s = 'tbparent') or (s = 'acparent') then
    begin
        if not SUCCEEDED(SDOM.get_parentNode(pEle)) then
            Exit;
    end
    else if (s = 'tbchild') or (s = 'acchild') then
    begin
        if not SUCCEEDED(SDOM.get_firstChild(pEle)) then
            Exit;
    end
    else if (s = 'tbprevs') or (s = 'acprevs') then
    begin
        if not SUCCEEDED(SDOM.get_previousSibling(pEle)) then
            Exit;
    end
    else if (s = 'tbnexts') or (s = 'acnexts') then
    begin
        try
            if not SUCCEEDED(SDOM.get_nextSibling(pEle)) then
                pEle := nil;
        except
            pEle := nil;
        end;
        try
            if pEle = nil then
            begin
              if SUCCEEDED(SDOM.get_parentNode(pEle)) then
              begin
                if SUCCEEDED(pEle.get_nodeInfo(PC, SI, PC2, PU, PU2, WD)) then
                begin
                  if Integer(WD) = 1 then
                  begin
                    if SUCCEEDED(pEle.get_firstChild(ppEle)) then
                    begin
                      if SUCCEEDED(ppEle.get_nodeInfo(PC, SI, PC2, PU, PU2, WD)) then
                      begin
                        if Integer(WD) = 1 then
                        begin
                          if SDOM = ppEle then
                            pEle := nil;
                        end;
                      end;
                    end;
                  end;
                end;
              end;
            end;
        except

        end;
    end
    else
        pEle := nil;
    iAcc := nil;
        if pEle <> nil then
        begin
            if SUCCEEDED(pEle.QueryInterface(IID_IServiceProvider, Serv)) then
            begin
                if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, iACC)) then
                begin
                    iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0);
                    if (acRect.Checked) then
                        ShowRectWnd2(clRed, RC);
                    //varParent := 0;
                    SDOM := pEle;
                    GetNaviState;
                    if (mnuMSAA.Checked <> false) and (Assigned(iACc)) then //MSAA text start
                    begin
                        MSAAText;
                    end;
                    if (mnuIA2.Checked) and (Assigned(iACc)) then
                            SetIA2Text;
                    if mnuARIA.Checked then
                        ARIAText4FF;
                    if mnuHTML.Checked then
                        HTMLText4FF;
                end;
            end;
        end;
end;

procedure TwndMSAAV.ExecMSAAMode(sender: TObject);
var
    pAcc: IAccessible;
    s: string;
    tc: TComponent;
    iDis: iDispatch;
    iRes, iDir: integer;
    ov:OleVariant;
    sNode: TTreeNode;
begin

    if (Sender is TComponent) then
    begin
        tc := Sender as TComponent;
    end
    else
        Exit;
    pAcc := nil;
    sNode := nil;
    s := LowerCase(tc.Name);
    if (s = 'tbparent') or (s = 'acparent') then
    begin
        try
        if TreeView1.Items.Count > 0 then
        begin
        if (TreeView1.SelectionCount > 0) then
        begin
            if TreeView1.Selected.Parent <> nil then
            begin
                sNode := TreeView1.Selected.Parent;
                if TTreeData(sNode.Data^).Acc <> nil then
                begin
                    pAcc := TTreeData(sNode.Data^).Acc;
                    varParent := TreeView_MapHTREEITEMtoAccID(TreeView1.Handle, sNode.ItemId); //TTreeData(sNode.Data^).iID;
                end;
            end;
        end;
        end
        else
        begin
            if (Assigned(iAcc))  then
            begin
                iRes := iAcc.Get_accParent(iDis);
                if (iRes = S_OK) and (iDis <> nil) then
                begin
                    varParent := 0;
                    if not SUCCEEDED(iDis.QueryInterface(IID_IACCESSIBLE, pACC)) then
                        pAcc := nil
                    else
                    begin

                        if SUCCEEDED(pAcc.Get_accParent(iDis)) then
                        accRoot := iDis as IAccessible;
                            accRoot := iDis as IAccessible;
                    end;
                end
                else
                    pAcc := nil;

            end;
        end;
        except
            on E:Exception do
                ShowErr(E.Message);
        end;
    end
    else if (s = 'tbchild') or (s = 'acchild') then
    begin
        if TreeView1.Items.Count > 0 then
        begin
            if (TreeView1.SelectionCount > 0) then
            begin
                if TreeView1.Selected.HasChildren then
                begin
                    sNode := TreeView1.Selected.getFirstChild;
                    if TTreeData(sNode.Data^).Acc <> nil then
                    begin
                        pAcc := TTreeData(sNode.Data^).Acc;
                        varParent := TreeView_MapHTREEITEMtoAccID(TreeView1.Handle, sNode.ItemId); //TTreeData(sNode.Data^).iID;
                    end;
                end;
            end;
        end
        else
        begin
            if  (Assigned(iAcc)) then
            begin
                iDir := NAVDIR_FIRSTCHILD;

                iRes := iAcc.accNavigate(iDir, varParent, ov);
                if (SUCCEEDED(iRes)) and (not VarIsType(ov, varNull)) then
                begin
                    if VarIsType(ov, varInteger) then
                    begin

                        pAcc := iAcc;//iDis as IACCESSIBLE;
                        varParent := TVarData(ov).VInteger;
                    end
                    else if VarIsType(ov, varDispatch) then
                    begin
                        pAcc := IAccessible(TVarData(ov).VDispatch);
                        varParent := 0;
                    end;
                    if SUCCEEDED(pAcc.Get_accParent(iDis)) then
                        accRoot := iDis as IAccessible;
                end;
            end;
        end;
    end
    else if  (s = 'tbprevs') or (s = 'acprevs') or (s = 'tbnexts') or (s = 'acnexts') then
    begin
        if TreeView1.Items.Count > 0 then
        begin
        if (TreeView1.SelectionCount > 0) then
        begin
            if (s = 'tbprevs') or (s = 'acprevs') then
            begin
                if TreeView1.Selected.getPrevSibling <> nil then
                begin
                    sNode := TreeView1.Selected.getPrevSibling;
                    if TTreeData(sNode.Data^).Acc <> nil then
                    begin
                        pAcc := TTreeData(sNode.Data^).Acc;
                        varParent := TreeView_MapHTREEITEMtoAccID(TreeView1.Handle, sNode.ItemId); //TTreeData(sNode.Data^).iID;
                    end;
                end;
            end
            else
            begin
                if TreeView1.Selected.getNextSibling <> nil then
                begin
                    sNode := TreeView1.Selected.getNextSibling;
                    if TTreeData(sNode.Data^).Acc <> nil then
                    begin
                        pAcc := TTreeData(sNode.Data^).Acc;
                        varParent := TreeView_MapHTREEITEMtoAccID(TreeView1.Handle, sNode.ItemId); //TTreeData(sNode.Data^).iID;
                    end;
                end;
            end;
        end;
        end
        else
        begin
        if  (Assigned(iAcc)) then
        begin
        if (s = 'tbprevs') or (s = 'acprevs') then
        begin
            iDir := NAVDIR_PREVIOUS;
        end
        else
            iDir := NAVDIR_NEXT;
        VariantInit(ov);
        iRes := iAcc.accNavigate(iDir, varParent, ov);
        if (SUCCEEDED(iRes))  then
        begin
            if VarIsType(ov, varInteger) then
            begin
                pAcc := accRoot;//iDis as IACCESSIBLE;
                varParent := TVarData(ov).VInteger;
            end
            else if (VarIsType(ov, varDispatch))  then
            begin

                pAcc := IAccessible(TVarData(ov).VDispatch);
                varParent := 0;
            end
            else
                pAcc := nil;
        end
        else
        begin

        end;
        end;
        end;
    end
    else
        pAcc := nil;
    try
    if pAcc <> nil {or (VarParent <> 0)} then
    begin
        iAcc := pAcc;
        if sNode <> nil then
        begin
            TreeMode := True;
            sNode.Selected := true;
        end
        else
        begin
            TreeMode := False;
            if (not acMSAAMode.Checked) then
            begin
              if iMode <> 2 then
              begin
                if (acRect.Checked) then
                begin
                  ShowRectWnd(clRed);
                end;
                ShowMSAAText;
              end
              else
                iAcc := nil;
            end
            else
            begin
              if (acRect.Checked) then
              begin
                ShowRectWnd(clRed);
              end;
              ShowMSAAText;
            end;
        end;

    end;
    except
        on E:Exception do
        begin
            ShowErr(E.Message);
        end;
    end;
end;

procedure TwndMSAAV.acNextSExecute(Sender: TObject);
begin
    TreeList1.Items.Clear;
    Memo1.Lines.Clear;
    ExecMSAAMode(acNextS);
end;

procedure TwndMSAAV.acPrevSExecute(Sender: TObject);
begin
    TreeList1.Items.Clear;
    Memo1.Lines.Clear;
    ExecMSAAMode(acPrevS);
end;

procedure TwndMSAAV.acOnlyFocusExecute(Sender: TObject);
begin
    acOnlyFOcus.Checked := not acOnlyFOcus.Checked;
    if acOnlyFocus.Checked then
        acMSAAMode.Checked := false;
    ExecOnlyFocus;
end;

procedure TwndMSAAV.acChildExecute(Sender: TObject);
begin
    TreeList1.Items.Clear;
    Memo1.Lines.Clear;
    ExecMSAAMode(acChild);
end;


procedure TwndMSAAV.acParentExecute(Sender: TObject);
begin
    TreeList1.Items.Clear;
    Memo1.Lines.Clear;
    ExecMSAAMode(acParent);
end;



procedure TwndMSAAV.acRectExecute(Sender: TObject);
begin
     acRect.Checked := not acRect.Checked;
    if acRect.Checked then
    begin
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        ShowRectWnd(clRed);
    end
    else
    begin
        FreeAndNil(WndFocus);
        FreeAndNil(WndLabel);
        FreeAndNil(WndDesc);
        FreeAndNil(WndTarg);
        WndFocus := nil;
        WndLabel := nil;
        WndDesc := nil;
        WndTarg := nil;
    end;
end;



procedure TwndMSAAV.acSettingExecute(Sender: TObject);
var
    WndSet: TWndSet;
    i, b:integer;

    ini: TMemINiFile;
begin
    WndSet := TWndSet.Create(self);
    try
        WndSet.Load_Str(TransPath);
        FormStyle := fsNormal;
        WndSet.cbExTip.Checked := Extip;
        WndSet.ChkList.Items.Add(lMSAA[0]);
        WndSet.ChkList.Header[WndSet.ChkList.Items.Count - 1] := True;
        for i := 1 to 11 do
        begin
            WndSet.ChkList.Items.Add(lMSAA[i]);
            b := Trunc(IntPower(2, i-1));
            if flgMSAA and b <> 0 then
            begin
                WndSet.chkList.Checked[i] := True;
            end;
        end;

        WndSet.ChkList.Items.Add(lIA2[0]);
        WndSet.ChkList.Header[WndSet.ChkList.Items.Count - 1] := True;
        for i := 1 to 10 do
        begin
            WndSet.ChkList.Items.Add(lIA2[i]);
            b := Trunc(IntPower(2, i-1));
            if flgIA2 and b <> 0 then
            begin
                WndSet.chkList.Checked[i + 12] := True;
            end;
        end;

        WndSet.ChkList.Items.Add(lUIA[0]);
        WndSet.ChkList.Header[WndSet.ChkList.Items.Count - 1] := True;
        for i := 1 to 54 do
        begin
            WndSet.ChkList.Items.Add(lUIA[i]);

            if i < 32 then
            begin
                b := Trunc(IntPower(2, i-1));
                if flgUIA and b <> 0 then
                begin
                    WndSet.chkList.Checked[i + 23] := True;
                end;
            end
            else
            begin
                b := Trunc(IntPower(2, i-32));
                if flgUIA2 and b <> 0 then
                begin
                    WndSet.chkList.Checked[i + 23] := True;
                end;
            end;
        end;



        if WndSet.ShowModal = mrOK then
        begin
            Font := WndSet.Font;
            flgMSAA := 0;
            for i := 0 to 10 do
            begin
                if WndSet.chkList.Checked[i + 1] then
                begin
                    b := Trunc(IntPower(2, i));
                    flgMSAA := flgMSAA or b;
                end;
            end;
            flgIA2 := 0;
            for i := 0 to 9 do
            begin
                if WndSet.chkList.Checked[i + 13] then
                begin
                    b := Trunc(IntPower(2, i));
                    flgIA2 := flgIA2 or b;
                end;
            end;
            flgUIA := 0;
            flgUIA2 := 0;
            for i := 0 to 53 do
            begin
                if WndSet.chkList.Checked[i + 24] then
                begin
                    if i < 31 then
                    begin
                      b := Trunc(IntPower(2, i));

                      flgUIA := flgUIA or b;
                    end
                    else
                    begin
                      b := Trunc(IntPower(2, i-31));
                      flgUIA2 := flgUIA2 or b;
                    end;
                end;
            end;
            if (flgMSAA = 0) and (flgIA2 = 0) and (flgUIA = 0) and (flgUIA2 = 0) then
                flgMSAA := 1;
            Extip := WndSet.cbExTip.Checked;
            ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
            try
                Ini.WriteString('Settings', 'FontName', Font.Name);
                ini.WriteInteger('Settings', 'FontSize', Font.Size);
                Ini.WriteInteger('Settings', 'Charset', Font.Charset);
                ini.WriteInteger('Settings', 'flgMSAA', flgMSAA);
                ini.WriteInteger('Settings', 'flgIA2', flgIA2);
                ini.WriteInteger('Settings', 'flgUIA', flgUIA);
                ini.WriteInteger('Settings', 'flgUIA2', flgUIA2);
                ini.WriteBool('Settings', 'ExTooltip', Extip);
                ini.UpdateFile;
            finally
                ini.Free;
            end;
        end;
    finally
        WndSet.Free;
        FormStyle := fsStayOnTop;
    end;
end;



procedure TwndMSAAV.acMMColExecute(Sender: TObject);
begin
    if Panel4.Height > 12 then
    begin
        Panel4.Height := 12;
        PB3.Picture.Bitmap.LoadFromResourceName(hInstance, 'UP');
        //Panel3.Align := alClient;
        PB2.Enabled := false;
        acTLCol.Enabled := False;
        Splitter2.Enabled := false;
    end
    else
    begin
        PB3.Picture.Bitmap.LoadFromResourceName(hInstance, 'DOWN');
        if P4H <= 12 then
            P4H := P4HDEF;
        Panel4.Height := P4H;
        //Panel3.Align := alBottom;
        PB2.Enabled := True;
        acTLCol.Enabled := True;
        Splitter2.Enabled := True;
    end;
end;

procedure TwndMSAAV.acTLColExecute(Sender: TObject);
begin
    if Panel3.Height > 12 then
    begin
        //Panel3.Height := 22;
        Panel4.Height := Panel5.Height - 12;
        PB2.Picture.Bitmap.LoadFromResourceName(hInstance, 'DOWN');
        //Panel4.Align := alClient;
        PB3.Enabled := false;
        acMMCol.Enabled := false;
        Splitter2.Enabled := false;
    end
    else
    begin
        PB2.Picture.Bitmap.LoadFromResourceName(hInstance, 'UP');
        if P4H <= 12 then
            P4H := P4HDEF;
        if P4H = Panel5.Height - 12 then
            P4H := Panel5.Height div 2;
        Panel4.Height := P4H;
        //Panel4.Align := alBottom;
        PB3.Enabled := True;
        Splitter2.Enabled := True;
        acMMCol.Enabled := True;
    end;
    //caption := inttostr(P4H) + '/' + inttostr(Panel4.Height - 12);
end;

procedure TwndMSAAV.acTVcolExecute(Sender: TObject);
begin
    if Panel2.Width > 12 then
    begin
        Panel2.Width := 12;
        PB1.Picture.Bitmap.LoadFromResourceName(hInstance, 'RIGHT');
        Splitter1.Enabled := false;
    end
    else
    begin
        PB1.Picture.Bitmap.LoadFromResourceName(hInstance, 'LEFT');
        if P2W <= 12 then
            P2W := P2WDEF;
        Panel2.Width := P2W;
        Splitter1.Enabled := True;
    end;
end;

procedure TwndMSAAV.GetNaviState(AllFalse: boolean = false);
var
    iDis: IDispatch;
    ich, iRes:integer;
    ov: OleVariant;
begin
    if AllFalse then
    begin
        tbParent.Enabled := false;
        tbChild.Enabled := false;
        tbPrevS.Enabled := false;
        tbNextS.Enabled := false;
        acParent.Enabled := tbParent.Enabled;
        acChild.Enabled := tbChild.Enabled;
        acPrevS.Enabled := tbPrevS.Enabled;
        acNextS.Enabled := tbNextS.Enabled;
    end
    else
    begin
        GetNaviState(True);
        if TreeView1.Items.Count > 0 then
        begin
            if (TreeView1.SelectionCount > 0) then
            begin
                if TreeView1.Selected.Parent <> nil then
                begin
                    tbParent.Enabled := true;
                end;
                if TreeView1.Selected.HasChildren then
                begin
                    tbChild.Enabled := true;
                end;
                if TreeView1.Selected.getNextSibling <> nil then
                begin
                    tbNextS.Enabled := true;
                end;
                if TreeView1.Selected.getPrevSibling<> nil then
                begin
                    tbPrevS.Enabled := true;
                end;
            end;

        end
        else
        begin
            if Assigned(iAcc) then
            begin


                iRes := iAcc.Get_accParent(iDis);
                if (SUCCEEDED(iRes)) and (Assigned(iDis)) then
                    tbParent.Enabled := true;

                iAcc.Get_accChildCount(ich);
                if ich > 0 then
                begin
                    tbChild.Enabled := true;
                end;
                iRes := iAcc.accNavigate(NAVDIR_PREVIOUS, VarParent, ov);
                if (SUCCEEDED(iRes)) then
                //ov := iAcc.accNavigate(NAVDIR_PREVIOUS, varParent);
                if (not VarIsType(ov, varNull)) then
                begin
                    if VarisType(ov, varInteger) or VarisType(ov, varDispatch) then
                    begin
                        if (not VarIsEmpty(ov)) and (not VarIsClear(ov)) then
                            tbPrevS.Enabled := true;
                    end;
                end;
                iRes := iAcc.accNavigate(NAVDIR_NEXT, VarParent, ov);
                if (SUCCEEDED(iRes)) then
                if (not VarIsType(ov, varNull)) then
                begin
                    if VarisType(ov, varInteger) or VarisType(ov, varDispatch) then
                    begin
                        if (not VarIsEmpty(ov)) and (not VarIsClear(ov)) then
                            tbNextS.Enabled := true;
                    end;
                end;

            end;
        end;
        acParent.Enabled := tbParent.Enabled;
        acChild.Enabled := tbChild.Enabled;
        acPrevS.Enabled := tbPrevS.Enabled;
        acNextS.Enabled := tbNextS.Enabled;

    end;
end;

procedure TwndMSAAV.ExecOnlyFocus;
begin

    mnuSelD.Enabled := not acOnlyFocus.Checked;
    tbCopy.Enabled := not acOnlyFocus.Checked;
    tbCursor.Enabled := not acOnlyFocus.Checked;
    tbFocus.Enabled := not acOnlyFocus.Checked;
    tbRectAngle.Enabled := not acOnlyFocus.Checked;
    tbShowTip.Enabled := not acOnlyFocus.Checked;


    if acOnlyFocus.Checked then
    begin
        GetNaviState(True);
        if Assigned(WndTip) then
            FreeandNil(WndTip);
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        if WndFocus.Visible then
            ShowRectWnd(clYellow);
    end
    else
    begin
        if not acRect.Checked then
        begin
            FreeAndNil(WndFocus);
            WndFocus := nil;
        end
        else
        begin
            if WndFocus.Visible then
                ShowRectWnd(clRed);
            GetNaviState;
        end;
        if acShowtIp.Checked then
        begin
            ShowTipWnd;
        end;
    end;
end;

procedure TwndMSAAV.mnuLangChildClick(Sender: TObject);
var
    i: integer;
    bChk: boolean;
    lf: string;
    ini: TMemIniFile;
begin
    if not (Sender is TMenuitem) then
        Exit;
    (Sender as TMenuitem).Checked := not (Sender as TMenuitem).Checked;
    if mnuLang.Count > 0 then
    begin
        bChk := False;
        for i := 0 to mnuLang.Count - 1 do
        begin
            if i > LangList.Count - 1 then
                Break;
            if mnuLang.Items[i].Checked then
            begin
                bChk := True;
                lf := LangList[i];
                Break;
            end;
        end;
        if not bChk then
        begin
            if (Sender is TMenuitem) then
            begin
                if (Sender as TMenuitem).MenuIndex > LangList.Count - 1 then
                begin
                    mnuLang.Items[0].Checked := True;
                    lf := LangList[0];
                end
                else
                begin
                    (Sender as TMenuitem).Checked := True;
                    lf := LangList[(Sender as TMenuitem).MenuIndex];
                end;
            end;
        end;
        if LowerCase(TransPath) <> LowerCase(TransDir + lf)  then
        begin
            TransPath := TransDir + lf;
            //showmessage(Transpath);
            LoadLang;
            ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
            try
                Ini.WriteString('Settings', 'FontName', Font.Name);
                ini.WriteInteger('Settings', 'FontSize', Font.Size);
                Ini.WriteInteger('Settings', 'Charset', Font.Charset);
                Ini.WriteString('Settings', 'LangFile', lf);
                ini.UpdateFile;
            finally
                ini.Free;
            end;
        end;
    end;
end;

function TwndMSAAV.LoadLang: string;
var
	//ini: TMemIniFile;
    d, Msg, UIA_fail: string;
    cList: TStringList;
    i: integer;
begin
    Result := 'CoCreateInstance failed.';
    if FileExists(Transpath) then
    begin
        //ini := TMemIniFile.Create(TransPath, TEncoding.Unicode);
        cList := TStringList.Create;
	    try
            //lMSAA: array[0..11] of string;
            //lIA2: array [0..10] of string;
            //lUIA: array [0..28] of string;
            lMSAA[0] := LoadTranslation('MSAA', 'MSAA', 'MS Active Accessibility');
            lMSAA[1] := LoadTranslation('MSAA', 'accName','accName');
            lMSAA[2] := LoadTranslation('MSAA', 'accRole','accRole');
            lMSAA[3] := LoadTranslation('MSAA', 'accState','accState');
            lMSAA[4] := LoadTranslation('MSAA', 'accDescription','accDescription');
            lMSAA[5] := LoadTranslation('MSAA', 'accDefaultAction','accDefaultAction');
            lMSAA[6] := LoadTranslation('MSAA', 'accValue','accValue');
            lMSAA[7] := LoadTranslation('MSAA', 'accParent','accParent');
            lMSAA[8] := LoadTranslation('MSAA', 'accChildCount','accChildCount');
            lMSAA[9] := LoadTranslation('MSAA', 'accHelp','accHelp');
            lMSAA[10] := LoadTranslation('MSAA', 'accHelpTopic','accHelpTopic');
            lMSAA[11] := LoadTranslation('MSAA', 'accKeyboardShortcut','accKeyboardShortcut');


            lIA2[0] := LoadTranslation('IA2', 'IA2', 'IAccessible2');
            lIA2[1] := LoadTranslation('IA2', 'Name','Name');
            lIA2[2] := LoadTranslation('IA2', 'Role','Role');
            lIA2[3] := LoadTranslation('IA2', 'States','States');
            lIA2[4] := LoadTranslation('IA2', 'Description','Description');
            lIA2[5] := LoadTranslation('IA2', 'Relations','Relations');
            //lIA2[6] := LoadTranslation('IA2', 'RelationTargets','Relation Targets');
            lIA2[6] := LoadTranslation('IA2', 'Attributes','Object Attributes');
            lIA2[7] := LoadTranslation('IA2', 'Value','Value');
            lIA2[8] := LoadTranslation('IA2', 'LocalizedExtendedRole','LocalizedExtendedRole');
            lIA2[9] := LoadTranslation('IA2', 'LocalizedExtendedStates','LocalizedExtendedStates');
            lIA2[10] := LoadTranslation('IA2', 'RangeValue','RangeValue');
            lIA2[11] := LoadTranslation('IA2', 'RV_Value','Value');
            lIA2[12] := LoadTranslation('IA2', 'RV_Minimum','Minimum');
            lIA2[13] := LoadTranslation('IA2', 'RV_Maximum','Maximum');


            lUIA[0] := LoadTranslation('UIA', 'UIA', 'UIAutomation');
            lUIA[1] := LoadTranslation('UIA', 'CurrentAcceleratorKey','CurrentAcceleratorKey');
            lUIA[2] := LoadTranslation('UIA', 'CurrentAccessKey','CurrentAccessKey');
            lUIA[3] := LoadTranslation('UIA', 'CurrentAriaProperties','CurrentAriaProperties');
            lUIA[4] := LoadTranslation('UIA', 'CurrentAriaRole','CurrentAriaRole');
            lUIA[5] := LoadTranslation('UIA', 'CurrentAutomationId','CurrentAutomationId');
            lUIA[6] := LoadTranslation('UIA', 'CurrentBoundingRectangle','CurrentBoundingRectangle');
            lUIA[7] := LoadTranslation('UIA', 'CurrentClassName','CurrentClassName');
            lUIA[8] := LoadTranslation('UIA', 'CurrentControlType','CurrentControlType');
            lUIA[9] := LoadTranslation('UIA', 'CurrentCulture','CurrentCulture');
            lUIA[10] := LoadTranslation('UIA', 'CurrentFrameworkId','CurrentFrameworkId');
            lUIA[11] := LoadTranslation('UIA', 'CurrentHasKeyboardFocus','CurrentHasKeyboardFocus');
            lUIA[12] := LoadTranslation('UIA', 'CurrentHelpText','CurrentHelpText');
            lUIA[13] := LoadTranslation('UIA', 'CurrentIsControlElement','CurrentIsControlElement');
            lUIA[14] := LoadTranslation('UIA', 'CurrentIsContentElement','CurrentIsContentElement');
            lUIA[15] := LoadTranslation('UIA', 'CurrentIsDataValidForForm','CurrentIsDataValidForForm');
            lUIA[16] := LoadTranslation('UIA', 'CurrentIsEnabled','CurrentIsEnabled');
            lUIA[17] := LoadTranslation('UIA', 'CurrentIsKeyboardFocusable','CurrentIsKeyboardFocusable');
            lUIA[18] := LoadTranslation('UIA', 'CurrentIsOffscreen','CurrentIsOffscreen');
            lUIA[19] := LoadTranslation('UIA', 'CurrentIsPassword','CurrentIsPassword');
            lUIA[20] := LoadTranslation('UIA', 'CurrentIsRequiredForForm','CurrentIsRequiredForForm');
            lUIA[21] := LoadTranslation('UIA', 'CurrentItemStatus','CurrentItemStatus');
            lUIA[22] := LoadTranslation('UIA', 'CurrentItemType','CurrentItemType');
            lUIA[23] := LoadTranslation('UIA', 'CurrentLocalizedControlType','CurrentLocalizedControlType');
            lUIA[24] := LoadTranslation('UIA', 'CurrentName','CurrentName');
            lUIA[25] := LoadTranslation('UIA', 'CurrentNativeWindowHandle','CurrentNativeWindowHandle');
            lUIA[26] := LoadTranslation('UIA', 'CurrentOrientation','CurrentOrientation');
            lUIA[27] := LoadTranslation('UIA', 'CurrentProcessId','CurrentProcessId');
            lUIA[28] := LoadTranslation('UIA', 'CurrentProviderDescription','CurrentProviderDescription');
            lUIA[29] := LoadTranslation('UIA', 'CurrentControllerFor','CurrentControllerFor');
            lUIA[30] := LoadTranslation('UIA', 'CurrentDescribedBy','CurrentDescribedBy');
            lUIA[31] := LoadTranslation('UIA', 'CurrentFlowsTo','CurrentFlowsTo');
            lUIA[32] := LoadTranslation('UIA', 'CurrentLabeledBy','CurrentLabeledBy');
            lUIA[33] := LoadTranslation('UIA', 'CurrentLiveSetting','CurrentLiveSetting');
            //Added 2014/06/23
            lUIA[34] := LoadTranslation('UIA', 'IsDockPatternAvailable','IsDockPatternAvailable');
            lUIA[35] := LoadTranslation('UIA', 'IsExpandCollapsePatternAvailable','IsExpandCollapsePatternAvailable');
            lUIA[36] := LoadTranslation('UIA', 'IsGridItemPatternAvailable','IsGridItemPatternAvailable');
            lUIA[37] := LoadTranslation('UIA', 'IsGridPatternAvailable','IsGridPatternAvailable');
            lUIA[38] := LoadTranslation('UIA', 'IsInvokePatternAvailable','IsInvokePatternAvailable');
            lUIA[39] := LoadTranslation('UIA', 'IsMultipleViewPatternAvailable','IsMultipleViewPatternAvailable');
            lUIA[40] := LoadTranslation('UIA', 'IsRangeValuePatternAvailable','IsRangeValuePatternAvailable');
            lUIA[41] := LoadTranslation('UIA', 'IsScrollPatternAvailable','IsScrollPatternAvailable');
            lUIA[42] := LoadTranslation('UIA', 'IsScrollItemPatternAvailable','IsScrollItemPatternAvailable');
            lUIA[43] := LoadTranslation('UIA', 'IsSelectionItemPatternAvailable','IsSelectionItemPatternAvailable');
            lUIA[44] := LoadTranslation('UIA', 'IsSelectionPatternAvailable','IsSelectionPatternAvailable');
            lUIA[45] := LoadTranslation('UIA', 'IsTablePatternAvailable','IsTablePatternAvailable');
            lUIA[46] := LoadTranslation('UIA', 'IsTableItemPatternAvailable','IsTableItemPatternAvailable');
            lUIA[47] := LoadTranslation('UIA', 'IsTextPatternAvailable','IsTextPatternAvailable');
            lUIA[48] := LoadTranslation('UIA', 'IsTogglePatternAvailable','IsTogglePatternAvailable');
            lUIA[49] := LoadTranslation('UIA', 'IsTransformPatternAvailable','IsTransformPatternAvailable');
            lUIA[50] := LoadTranslation('UIA', 'IsValuePatternAvailable','IsValuePatternAvailable');
            lUIA[51] := LoadTranslation('UIA', 'IsWindowPatternAvailable','IsWindowPatternAvailable');
            lUIA[52] := LoadTranslation('UIA', 'IsItemContainerPatternAvailable','IsItemContainerPatternAvailable');
            lUIA[53] := LoadTranslation('UIA', 'IsVirtualizedItemPatternAvailable','IsVirtualizedItemPatternAvailable');
            //Added 2014/06/25
            lUIA[54] := LoadTranslation('UIA', 'RangeValue','RangeValue');
            lUIA[55] := LoadTranslation('UIA', 'RV_Value','Value');
            lUIA[56] := LoadTranslation('UIA', 'RV_IsReadOnly','IsReadOnly');
            lUIA[57] := LoadTranslation('UIA', 'RV_Minimum','Minimum');
            lUIA[58] := LoadTranslation('UIA', 'RV_Maximum','Maximum');
            lUIA[59] := LoadTranslation('UIA', 'RV_LargeChange','LargeChange');
            lUIA[60] := LoadTranslation('UIA', 'RV_SmallChange','SmallChange');



            HelpURL := LoadTranslation('General', 'Help_URL', 'http://www.google.com/');
            Caption := LoadTranslation('General', 'MSAA_Caption', 'Accessibility Viewer');
            Font.Name := LoadTranslation('General', 'FontName', 'Arial');
            Font.Size := LoadTransInt('General', 'FontSize', 9);
            d := LoadTranslation('General', 'Charset', 'ASCII_CHARSET');
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
            mnuSelD.Caption := LoadTranslation('General', 'MSAA_gbSelDisplay', 'Select Display');
            tbFocus.Hint := LoadTranslation('General', 'MSAA_tbFocusHint', 'Watch Focus');
            tbCursor.Hint := LoadTranslation('General', 'MSAA_tbCursorHint', 'Watch Cursor');
            tbRectAngle.Hint := LoadTranslation('General', 'MSAA_tbRectAngleHint', 'Show Highlight Rectangle');

            tbShowtip.Hint :=  LoadTranslation('General', 'MSAA_tbBalloonHint', 'Show Balloon tip');
            acShowTip.Hint := tbShowtip.Hint;
            tbCopy.Hint := LoadTranslation('General', 'MSAA_tbCopyHint', 'Copy Text to Clipborad');
            tbOnlyFocus.Hint := LoadTranslation('General', 'MSAA_tbFocusOnly', 'Focus rectangle only');

            tbParent.Hint := LoadTranslation('General', 'MSAA_tbParentHint', 'Navigates to parent object');
            tbChild.Hint := LoadTranslation('General', 'MSAA_tbChildHint', 'Navigates to first child object');
            tbPrevS.Hint := LoadTranslation('General', 'MSAA_tbPrevSHint', 'Navigates to previous sibling object');
            tbNextS.Hint := LoadTranslation('General', 'MSAA_tbNextSHint', 'Navigates to next sibling object');
            tbHelp.Hint := LoadTranslation('General', 'MSAA_tbHelpHint', 'Show online help');
            tbHelp.Hint := tbHelp.Hint + '(' +HelpURL + ')';

            //tbFont.Hint := LoadTranslation('Generals', 'MSAA_tbFontHint', 'Show font dialog');
            tbMSAAMode.Hint := LoadTranslation('General', 'MSAA_tbMSAAModeHint', '');
            tbRegister.Hint := LoadTranslation('General', 'MSAA_tbMSAASetHint', 'Show Setting Dialog');
            //acFocus.Hint := LoadTranslation('General', 'MSAA_tbFocusHint', 'Watch Focus(F9)');
            //acCursor.Hint := LoadTranslation('General', 'MSAA_tbCursorHint', 'Watch Cursor(F10)');
            //acRect.Hint := LoadTranslation('General', 'MSAA_tbRectAngleHint', 'Show Highlight Rectangle(F11)');
            //acCopy.Hint := LoadTranslation('General', 'MSAA_tbCopyHint', 'Copy Text to Clipborad(F12)');
            mnuLang.Caption := LoadTranslation('General', 'mnuLang', '&Language');
            mnuView.Caption := LoadTranslation('General', 'mnuView', '&View');
            mnuMSAA.Caption := LoadTranslation('General', 'MSAA_cbMSAA', 'MSAA');
            mnuARIA.Caption := LoadTranslation('General', 'MSAA_cbARIA', 'ARIA');
            mnuHTML.Caption := LoadTranslation('General', 'MSAA_cbHTML', 'HTML');
            mnuIA2.Caption := LoadTranslation('General', 'MSAA_cbIA2', 'IA2');
            mnuUIA.Caption := LoadTranslation('General', 'MSAA_cbUIA', 'UIAutomation');
            mnuMSAA.Hint := LoadTranslation('General', 'MSAA_cbMSAAHint', '');
            mnuARIA.Hint := LoadTranslation('General', 'MSAA_cbARIAHint', '');
            mnuHTML.Hint := LoadTranslation('General', 'MSAA_cbHTMLHint', '');
            mnuIA2.Hint := LoadTranslation('General', 'MSAA_cbIA2Hint', '');
            mnuUIA.Hint := LoadTranslation('General', 'MSAA_cbUIAHint', '');
            None := LoadTranslation('General', 'none', '(none)');
            ConvErr := LoadTranslation('General', 'MSAA_ConvertError', 'Format is invalid: %s');
            Msg := LoadTranslation('General', 'MSAA_HookIsFailed', 'SetWinEventHook is Failed!!');

            sTrue := LoadTranslation('General', 'true', 'True');
            sFalse := LoadTranslation('General', 'false', 'False');
            mnuTVSave.Caption := LoadTranslation('General', 'mnuSave', '&Save');
            mnuTVSAll.Caption := LoadTranslation('General', 'mnuSaveAll', '&All items');
            mnuTVSSel.Caption := LoadTranslation('General', 'mnuSaveSelect', '&Selected items');

            mnuSAll.Caption := LoadTranslation('General', 'mnuTVSAll', 'Save &All items');
            mnuSSel.Caption := LoadTranslation('General', 'mnuTVSSelect', 'Save &Selected items');



            tbFocus.Caption := LoadTranslation('General', 'MSAA_tbFocusName', 'Focus');
            tbCursor.Caption := LoadTranslation('General', 'MSAA_tbCursorName', 'Cursor');
            tbRectAngle.Caption := LoadTranslation('General', 'MSAA_tbRectAngleName', 'Highlight Rectangle');
            tbShowtip.Caption :=  LoadTranslation('General', 'MSAA_tbBalloonName', 'Balloon tip');
            tbCopy.Caption := LoadTranslation('General', 'MSAA_tbCopyName', 'Copy');
            tbOnlyFocus.Caption := LoadTranslation('General', 'MSAA_tbFocusOnlyName', 'Focus only');
            tbParent.Caption := LoadTranslation('General', 'MSAA_tbParentName', 'Parent');
            tbChild.Caption := LoadTranslation('General', 'MSAA_tbChildName', 'Child');
            tbPrevS.Caption := LoadTranslation('General', 'MSAA_tbPrevSName', 'Previous');
            tbNextS.Caption := LoadTranslation('General', 'MSAA_tbNextSName', 'Next');
            tbHelp.Caption := LoadTranslation('General', 'MSAA_tbHelpName', 'Help');
            //tbFont.Caption := LoadTranslation('General', 'MSAA_tbFontName', 'Font');
            tbMSAAMode.Caption := LoadTranslation('General', 'MSAA_tbMSAAModeName', 'MSAA');
            tbRegister.Caption := LoadTranslation('General', 'MSAA_tbMSAASetName', 'Settings');

            mnuReg.Caption := LoadTranslation('General', 'MSAA_mnuReg', '&Register');
            mnuUnReg.Caption := LoadTranslation('General', 'MSAA_mnuUnreg', '&Unregister');
            mnuReg.Hint := LoadTranslation('General', 'MSAA_mnuRegHint', 'Register IAccessible2Proxy.dll');
            mnuUnReg.Hint := LoadTranslation('General', 'MSAA_mnuUnregHint', 'Unregister IAccessible2Proxy.dll');

            Memo1.AccName := LoadTranslation('General', 'MSAA_CodeEdit', 'Code');
            Memo1.Hint := LoadTranslation('General', 'MSAA_CodeEditHint', '');
            Memo1.AccDesc := GetLongHint(Memo1.Hint);
            TreeList1.Columns[0].Header := LoadTranslation('General', 'MSAA_tlColumn1', 'Intarfece');
            TreeList1.Columns[1].Header := LoadTranslation('General', 'MSAA_tlColumn2', 'Value');
            TreeList1.AccName := PChar(LoadTranslation('General', 'MSAA_ListName', 'Result List'));
            d := PChar(LoadTranslation('General', 'MSAA_ListName', 'Result List'));
            SetWindowText(TreeList1.Handle, PWideChar(d));
            TreeList1.Hint := LoadTranslation('General', 'MSAA_ListHint', '');
            TreeList1.AccDesc := GetLongHint(TreeList1.Hint);
            TreeView1.AccName := PChar(LoadTranslation('General', 'MSAA_TVName', 'Accessibility Tree'));
            //d := PChar(LoadTranslation('General', 'MSAA_TVName', 'Accessibility Tree'));
            //SetWindowText(TreeView1.Handle, PWideChar(d));
            TreeView1.Hint := LoadTranslation('General', 'MSAA_TVHint', '');
            TreeView1.AccDesc := GetLongHint(TreeView1.Hint);
            UIA_fail := LoadTranslation('General', 'UIA_CreationError', 'CoCreateInstance failed. ');

            PB1.Hint :=  LoadTranslation('General', 'Collapse_TV_Hint', 'Collapse(or Expand) Button for TreeView');
            PB2.Hint :=  LoadTranslation('General', 'Collapse_TL_Hint', 'Collapse(or Expand) Button for TreeList');
            PB3.Hint :=  LoadTranslation('General', 'Collapse_CE_Hint', 'Collapse(or Expand) Button for CodeEdit');
            PB1.AccDesc := PB1.Hint;
            PB2.AccDesc := PB2.Hint;
            PB3.AccDesc := PB3.Hint;
            acTVCol.Caption := LoadTranslation('General', 'mnuTreeView', '&TreeView');
            acTLCol.Caption := LoadTranslation('General', 'mnuTreeList', 'Tree&List');
            acMMCol.Caption := LoadTranslation('General', 'mnuCodeEdit', '&CodeEdit');
            mnuColl.Caption := LoadTranslation('General', 'mnuCollapse', '&Collapse');

            mnuTVcont.Caption := LoadTranslation('General', 'mnuTVContent', '&Treeview contents');
            mnuTarget.Caption := LoadTranslation('General', 'mnuTarget', '&Target only');
            mnuAll.Caption := LoadTranslation('General', 'mnuAll', '&All related objects');

            rType := LoadTranslation('IA2', 'RelationType', 'Relation Type');
            rTarg := LoadTranslation('IA2', 'RelationTargets', 'Relation Targets');
            mnuBln.Caption := LoadTranslation('General', 'mnubln', 'Balloon tip(&B)');
            mnublnMSAA.Caption :=  LoadTranslation('General', 'mnublnMSAA', '&MS Active Accessibility');
            mnublnIA2.Caption :=  LoadTranslation('General', 'mnublnIA2', '&IAccessible2');
            mnublnCode.Caption :=  LoadTranslation('General', 'mnublnCode', '&Source code');

            ShowSrcLen := LoadTransInt('General', 'ShowSrcLen', 2000);
            cList.CommaText := LoadTranslation('General', 'ClassNames', '"mozillawindowclass","chrome_renderwidgethosthwnd","mozillawindowclass","chrome_widgetwin_0","chrome_widgetwin_1"');
            SetLength(CNames, cList.Count);
            for i := 0 to cList.Count - 1 do
            begin
              CNames[i] := cList.Strings[i];
            end;

            SaveDlg.Filter :=  LoadTranslation('General', 'SaveDLGFilter', 'HTML File|*.htm*|Text File|*.txt|All|*.*');
        finally
            Result := UIA_fail;
            cList.Free;
        end;
    end
    else
        TransPath := '';
end;

procedure TwndMSAAV.Load;
var
	ini: TMemIniFile;
    d, Msg, UIA_fail: string;
    b:boolean;
    ic: TIcon;
    hLib: Thandle;
    i: integer;
    mItem: TMenuItem;
    hr: HResult;
begin
    hHook := SetWinEventHook(EVENT_MIN,EVENT_MAX, 0, @WinEventProc, 0, 0,
        WINEVENT_OUTOFCONTEXT or WINEVENT_SKIPOWNPROCESS );
    Msg := 'SetWinEventHook is Failed!!';
    //flgMSAA, flgIA2, flgUIA: integer;
    flgMSAA := 124;
    flgIA2 := 119;
    flgUIA := 50332680;
    flgUIA2 := 0;
    iDefIndex := -1;


    //iAccS := nil;
    ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
    try
        if mnuLang.Visible then
        begin
            d := Ini.ReadString('Settings', 'LangFile', 'Default.ini');
            TransPath := TransDir + d;
        end;

        UIA_fail := LoadLang;
        Width := ini.ReadInteger('Settings', 'Width', 335);
        Height := ini.ReadInteger('Settings', 'Height', 500);
        //P2W, P4H: integer;
        Panel2.Width := ini.ReadInteger('Settings', 'P2W', 350);
        Panel4.Height := ini.ReadInteger('Settings', 'P4H', 250);

        flgMSAA := ini.ReadInteger('Settings', 'flgMSAA', flgMSAA);
        flgIA2 := ini.ReadInteger('Settings', 'flgIA2', flgIA2);
        flgUIA := ini.ReadInteger('Settings', 'flgUIA', flgUIA);
        flgUIA2 := ini.ReadInteger('Settings', 'flgUIA2', flgUIA2);
        if (flgMSAA = 0) and (flgIA2 = 0) and (flgUIA = 0) and (flgUIA2 = 0) then
            flgMSAA := 1;
        ExTip := ini.ReadBool('Settings', 'ExTooltip', False);
        Top := ini.ReadInteger('Settings', 'Top', (Screen.Height - Height) div 2);
        Left := ini.ReadInteger('Settings', 'Left', (Screen.Width - Width) div 2);
        Font.Name := Ini.ReadString('Settings', 'FontName', Font.Name);
            Font.Size := ini.ReadInteger('Settings', 'FontSize', Font.Size);
            Font.Charset := ini.ReadInteger('Settings', 'Charset', 0);


        mnuMSAA.Checked := ini.ReadBool('Settings', 'vMSAA', True);
        mnuARIA.Checked := ini.ReadBool('Settings', 'vARIA', True);
        mnuHTML.Checked := ini.ReadBool('Settings', 'vHTML', True);
        mnuIA2.Checked := ini.ReadBool('Settings', 'vIA2', True);
        mnuUIA.Checked := ini.ReadBool('Settings', 'vUIA', True);

        mnublnMSAA.Checked := ini.ReadBool('Settings', 'bMSAA', True);
        mnublnIA2.Checked := ini.ReadBool('Settings', 'bIA2', True);
        mnublnCode.Checked := ini.ReadBool('Settings', 'bCode', True);

        b := ini.ReadBool('Settings', 'TVAll', False);
        mnuAll.Checked := b;
        mnuTarget.Checked := not b;
        if (not mnuMSAA.Checked) and (not mnuARIA.Checked) and (not mnuHTML.Checked) and (not mnuIA2.Checked) and (not mnuUIA.Checked) then
            mnuMSAA.Checked := True;
    finally
        ini.Free;
    end;
    if mnuLang.Visible then
    begin
    if LangList.Count = 0 then
        mnuLang.Enabled := False;
    for i := 0 to LangList.Count - 1 do
    begin
        if FileExists(TransDir + LangList[i]) then
        begin
            //ini := TMemIniFile.Create(TransDir + LangList[i], TEncoding.Unicode);
            try
                mItem := MainMenu1.CreateMenuItem;
                mItem.Caption := LoadTranslation_Path('General', 'Language', 'English', TransDir + LangList[i]);
                mItem.RadioItem := True;
                mItem.GroupIndex := 1;
                //mItem.AutoCheck := True;
                mItem.OnClick := mnuLangChildClick;
                if LowerCase(TransPath) = LowerCase(TransDir + LangList[i])  then
                    mItem.Checked := True
                else
                    mItem.Checked := False;
                mnuLang.Add(mItem);
            finally
                //ini.Free;
            end;
        end;
    end;
    if LangList.Count = 1 then
        mnuLang.Items[0].Checked := True;
    end;
    SizeChange;

    //PB2.Picture.Bitmap.LoadFromResourceName(hInstance, 'UP');

    if Panel2.Width > 12 then
    begin
        P2W := Panel2.Width;
        PB1.Picture.Bitmap.LoadFromResourceName(hInstance, 'LEFT');
        Splitter1.Enabled := True;
    end
    else
    begin
        P2W := P2WDEF;
        PB1.Picture.Bitmap.LoadFromResourceName(hInstance, 'RIGHT');
        Splitter1.Enabled := false;
    end;
    if Panel4.Height > 12 then
    begin
        P4H := Panel4.Height;
        PB3.Picture.Bitmap.LoadFromResourceName(hInstance, 'DOWN');
        PB2.Enabled := True;
        acTLCol.Enabled := True;
        //Splitter2.Enabled := True;
    end
    else
    begin
        P4H := P4HDEF;
        PB3.Picture.Bitmap.LoadFromResourceName(hInstance, 'UP');
        PB2.Enabled := false;
        acTLCol.Enabled := false;
        //Splitter2.Enabled := false;
    end;
    if Panel3.Height > 12 then
    begin
        //P4H := Panel4.Height;
        PB2.Picture.Bitmap.LoadFromResourceName(hInstance, 'UP');
        PB3.Enabled := True;
        acMMCol.Enabled := True;
        //Splitter2.Enabled := True;
    end
    else
    begin
        P4H := P4HDEF;
        PB2.Picture.Bitmap.LoadFromResourceName(hInstance, 'DOWN');
        PB3.Enabled := false;
        acMMCol.Enabled := false;
        //Splitter2.Enabled := false;
    end;
    if (Panel3.Height <=12) or (Panel4.Height <= 12) then
        Splitter2.Enabled := false
    else
        Splitter2.Enabled := True;
    //CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
    CoInitialize(nil);
    {UIAuto := TCUIAutomation.Create(Self);
    if not Assigned(UIAuto) then
    begin
        Showmessage(UIA_fail);
        //UIA := nil;
        UIAuto := nil;
        mnuView.Enabled := False;
        Toolbar1.Enabled := False;
        Panel1.Enabled := False;
        Panel2.Enabled := False;
        Panel5.Enabled := False;
        //mnuUIA.Checked := false;
        //mnuUIA.Enabled := false;

    end;  }
    hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER, IID_IUIAutomation, UIAuto);

    if (UIAuto = nil) or (hr <> S_OK) then
    begin
        Showmessage(UIA_fail);

        UIAuto := nil;
        mnuView.Enabled := False;
        Toolbar1.Enabled := False;
        Panel1.Enabled := False;
        Panel2.Enabled := False;
        Panel5.Enabled := False;
        //mnuUIA.Checked := false;
        //mnuUIA.Enabled := false;

    end;
    TreeList1.Columns[0].Width := TreeList1.ClientWidth div 2;
    TreeList1.Columns[1].Width := TreeList1.ClientWidth div 2;
    if hHook = 0 then
    begin
        ShowMessage(Msg);
        TreeList1.Enabled := false;
        Toolbar1.Enabled := false;
        mnuSelD.Enabled := false;
    end;
    if IsWinVista then
    begin
        try
            ImageList2.Handle := ImageList_Create(16, 16, ILC_COLOR32, 0, Imagelist2.AllocBy);
        except

        end;
        hLib := LoadLibrary('user32');
        if hLib <> 0 then
        begin
            ic := TIcon.Create;
            try
                ic.Handle := LoadImage(hLib, MAKEINTRESOURCE(106), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR );
                ImageList2.AddIcon(ic);
                mnuReg.ImageIndex := 0;
                mnuUnReg.ImageIndex := 0;
            finally
                FreeLibrary(hLib);
                ic.Free;
            end;
        end;
    end;
    if not FileExists(DllPath) then
    begin
        mnuReg.Enabled := False;
        mnuUnreg.Enabled := False;
    end;

    Toolbar1.Focused;
end;



procedure TwndMSAAV.mnuMSAAClick(Sender: TObject);
begin
    if (not mnuMSAA.Checked) and (not mnuARIA.Checked) and (not mnuHTML.Checked) and (not mnuIA2.Checked) and (not mnuUIA.Checked) then
            mnuMSAA.Checked := True;
end;



procedure TwndMSAAV.mnuAllClick(Sender: TObject);
begin
    if (not mnuTarget.Checked) and (not mnuAll.Checked) then
        mnuAll.Checked := True;
end;

procedure TwndMSAAV.mnuTargetClick(Sender: TObject);
begin
    if (not mnuTarget.Checked) and (not mnuAll.Checked) then
        mnuTarget.Checked := True;
end;

procedure TwndMSAAV.ReflexTV(cNode: TTreeNode; var HTML: string; var iCnt: integer; ForSel: boolean = false);
var
	i, ulCnt: integer;
  tab, temp: string;
  Role: string;
  ws: widestring;
  ovChild: OleVariant;
  iRole :integer;
  ovValue: OleVariant;

  function GetLIContents(Acc: IAccessible; Child: integer; pNode: TTreeNode ): string;
  var
  	Res: string;
    PC:PChar;
    isp: iserviceprovider;
  begin
  	Role:= '';
    ovChild := Child;
    Acc.Get_accRole(ovChild, ovValue);
    iRole := TVarData(ovValue).VInteger;
    try
    if VarHaveValue(ovValue) and VarIsNumeric(ovValue) then
    begin

        if pNode.Text = '' then
        begin
        	Acc.Get_accName(ovChild, ws);
          if ws = '' then ws := None;
          ws := ReplaceStr(ws, '<', '&lt;');
          ws := ReplaceStr(ws, '>', '&gt;');
         	PC := StrAlloc(255);
          GetRoleTextW(ovValue, PC, StrBufSize(PC));
          Role := PC;
          StrDispose(PC);
          pNode.Text := ws + ' - ' + Role;
        end;
    end;
    if (iRole = ROLE_SYSTEM_STATICTEXT) or (iRole = ROLE_SYSTEM_TEXT) or (iRole = ROLE_SYSTEM_APPLICATION)  or (iRole = ROLE_SYSTEM_WINDOW) then
  	begin
    	ws := ReplaceStr(pNode.Text, '<', '&lt;');
      ws := ReplaceStr(ws, '>', '&gt;');
  		Res := Res + #13#10 + tab + '<li>' + ws;
  	end
  	else
  	begin

    	try
  		Res := Res + #13#10 + tab + '<li>' + #13#10 + '<input type="checkbox" id="disclosure' + inttostr(iCnt) + '" title="check to display details below" aria-controls="x-details' + inttostr(iCnt) + '"> ';
    	Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt) + '">' + pNode.Text + '</label> ';
    	Res := Res + #13#10 + tab + '<section id="x-details' + inttostr(iCnt) + '">';

      Res := Res + MSAAText4HTML(Acc, tab);

      if (iMode = 1) and (SUCCEEDED(Acc.QueryInterface(IID_ISERVICEPROVIDER, isp))) then//FF
      begin
      	Res := Res + IA2Text4HTML(Acc, tab);
      end;
      finally
      	Res := Res + #13#10 + tab + '</section>';
      	Inc(iCnt);
      end;

  	end;

    finally
    	Result := Res;
    end;
  end;

begin
	tab := #9;
  ulCnt := 0;
	if cNode.Level > 0 then
  begin
		for i := 1 to cNode.Level do
  		tab := tab + #9;
  end;
  if ForSel then
  begin
  	if cNode.Selected then
    begin
			HTML := HTML + GetLIContents(TTreeData(cNode.Data^).Acc, TTreeData(cNode.Data^).iID, cNode);
      Inc(ulCnt);
    end;
  end
  else
  	HTML := HTML + GetLIContents(TTreeData(cNode.Data^).Acc, TTreeData(cNode.Data^).iID, cNode);

  if cNode.HasChildren then
  begin
  	temp := '';
  	//HTML := HTML + #13#10 + tab + '<ul>';
    for i := 0 to cNode.Count - 1 do
    begin
    	Application.ProcessMessages;

      if cNode.Item[i].HasChildren then
      begin
        if ForSel then
  			begin
  				if cNode.Item[i].Selected then
       			temp := temp + #13#10;
        end
  			else
        	temp := temp + #13#10;
				ReflexTV(cNode.Item[i], temp, iCnt, Forsel);
      end
      else
      begin
      	if ForSel then
  			begin
  				if cNode.Item[i].Selected then
          begin
       			temp := temp + #13#10 + tab + #9 + GetLIContents(TTreeData(cNode.Item[i].Data^).Acc, TTreeData(cNode.Item[i].Data^).iID, cNode.Item[i]) + '</li>';
            Inc(ulCnt);
          end;
        end
  			else
        	temp := temp + #13#10 + tab + #9 + GetLIContents(TTreeData(cNode.Item[i].Data^).Acc, TTreeData(cNode.Item[i].Data^).iID, cNode.Item[i]) + '</li>';
      end;
    end;
    if temp <> '' then
    begin
      if ForSel then
      begin //ulの回数が多い？
        HTML := HTML + #13#10 + tab + '<ul>' + temp + #13#10 + tab + '</ul>';
      end
      else
        HTML := HTML + #13#10 + tab + '<ul>' + temp + #13#10 + tab + '</ul>';

    //HTML := HTML + #13#10 + tab + '</ul>';
    end;
  end;
  if ForSel then
  begin
  	if cNode.Selected then
			HTML := HTML  + '</li>';
  end
  else
  	HTML := HTML  + '</li>';

end;

procedure TwndMSAAV.mnuSAllClick(Sender: TObject);
begin
	mnuTVSAllClick(self);
end;

procedure TwndMSAAV.mnuSSelClick(Sender: TObject);
begin
	mnuTVSSelClick(self);
end;

procedure TwndMSAAV.mnuTVSAllClick(Sender: TObject);
var
	d, temp: string;
  sList: TStringList;
  i: integer;
begin
		//Save Treeview contents
    if TreeView1.Items.Count > 0 then
    begin
      i := 0;

    	ReflexTV(TreeView1.Items.Item[0], d, i);
      sList := TStringList.Create;
      try
      	temp := APPDir + 'output.html';
        if FileExists(temp) then
        begin
        	sList.LoadFromFile(temp, TEncoding.UTF8);
          sList.Text := StringReplace(sList.Text, '%contents%', d, [rfReplaceAll, rfIgnoreCase]);
        end
        else
        	sList.Text := d;

        if SaveDLG.Execute then
        begin
        	sList.SaveToFile(SaveDLG.FileName, TEncoding.UTF8);
        end;

      finally
        sList.Free;
      end;
    end;

end;

procedure TwndMSAAV.mnuTVSSelClick(Sender: TObject);
var
	d, temp: string;
  sList: TStringList;
  //i: integer;
  	i,iCnt: integer;
  tab{, temp}: string;
  Role: string;
  ws: widestring;
  ovChild: OleVariant;
  iRole :integer;
  ovValue: OleVariant;
  function GetLIContents(Acc: IAccessible; Child: integer; pNode: TTreeNode ): string;
  var
  	Res: string;
    PC:PChar;
    isp: iserviceprovider;
  begin
  	Role:= '';
    ovChild := Child;
    Acc.Get_accRole(ovChild, ovValue);
    iRole := TVarData(ovValue).VInteger;
    try
    if VarHaveValue(ovValue) and VarIsNumeric(ovValue) then
    begin

        if pNode.Text = '' then
        begin
        	Acc.Get_accName(ovChild, ws);
          if ws = '' then ws := None;
          ws := ReplaceStr(ws, '<', '&lt;');
          ws := ReplaceStr(ws, '>', '&gt;');
         	PC := StrAlloc(255);
          GetRoleTextW(ovValue, PC, StrBufSize(PC));
          Role := PC;
          StrDispose(PC);
          pNode.Text := ws + ' - ' + Role;
        end;
    end;
    if (iRole = ROLE_SYSTEM_STATICTEXT) or (iRole = ROLE_SYSTEM_TEXT) or (iRole = ROLE_SYSTEM_APPLICATION)  or (iRole = ROLE_SYSTEM_WINDOW) then
  	begin
    	ws := ReplaceStr(pNode.Text, '<', '&lt;');
      ws := ReplaceStr(ws, '>', '&gt;');
  		Res := Res + #13#10 + tab + '<li>' + ws;
  	end
  	else
  	begin

    	try
  		Res := Res + #13#10 + tab + '<li>' + #13#10 + '<input type="checkbox" id="disclosure' + inttostr(iCnt) + '" title="check to display details below" aria-controls="x-details' + inttostr(iCnt) + '"> ';
    	Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt) + '">' + pNode.Text + '</label> ';
    	Res := Res + #13#10 + tab + '<section id="x-details' + inttostr(iCnt) + '">';

      Res := Res + MSAAText4HTML(Acc, tab);
      {IA2[0] := LoadTranslation('IA2', 'IA2', 'IAccessible2');
            lIA2[1] := LoadTranslation('IA2', 'Name','Name');
            lIA2[2] := LoadTranslation('IA2', 'Role','Role');
            lIA2[3] := LoadTranslation('IA2', 'States','States');
            lIA2[4] := LoadTranslation('IA2', 'Description','Description');}
      if (iMode = 1) and (SUCCEEDED(Acc.QueryInterface(IID_ISERVICEPROVIDER, isp))) then//FF
      begin
      	Res := Res + IA2Text4HTML(Acc, tab);
      end;
      finally
      	Res := Res + #13#10 + tab + '</section>';
      	Inc(iCnt);
      end;

  	end;

    finally
    	Result := Res;
    end;
  end;
begin
		//Save Treeview contents
    if (TreeView1.SelectionCount > 0) and (TreeView1.Items.Count > 0) then
    begin
      //showmessage(inttostr(Treeview1.SelectionCount));
      {i := 0;
    	ReflexTV(TreeView1.Items.Item[0], d, i, True);  }
      iCnt := 0;
      for i := 0 to TreeView1.SelectionCount - 1 do
      begin
        d := d + GetLIContents(TTreeData(TreeView1.Selections[i].Data^).Acc, TTreeData(TreeView1.Selections[i].Data^).iID, TreeView1.Selections[i]) + '</li>' + #13#10;
      end;
      d := '<ul>' + d + '</ul>';
      sList := TStringList.Create;
      try
      	temp := APPDir + 'output.html';
        if FileExists(temp) then
        begin
        	sList.LoadFromFile(temp, TEncoding.UTF8);
          sList.Text := StringReplace(sList.Text, '%contents%', d, [rfReplaceAll, rfIgnoreCase]);
        end
        else
        	sList.Text := d;
        if SaveDLG.Execute then
        begin
        	sList.SaveToFile(SaveDLG.FileName, TEncoding.UTF8);
        end;

      finally
        sList.Free;
      end;
    end;

end;

procedure TwndMSAAV.SizeChange;
var
    SZ: TSize;
begin

    GetTextExtentPoint32W(Canvas.Handle, PWideChar(Caption), Length(PwideChar(Caption)), SZ);
    TreeList1.Font := Font;
    TreeList1.HeaderSettings.Font := Font;
    TreeList1.HeaderSettings.Height := SZ.Height + 3;
    TreeList1.Columns[0].Font := Font;
    TreeList1.Columns[1].Font := Font;

end;

procedure TwndMSAAV.Splitter1Moved(Sender: TObject);
begin
    P2W := Panel2.Width;
end;

procedure TwndMSAAV.Splitter2Moved(Sender: TObject);
begin
    P4H := Panel4.Height;
end;

function TwndMSAAV.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;


function TwndMSAAV.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  Result := E_NOTIMPL;
  pointer(TypeInfo) := nil;
end;

function TwndMSAAV.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := E_NOTIMPL;
  Count := 0;
end;

procedure BuildPositionalDispIds(pDispIds: PDispIdList; const dps: TDispParams);
var
  i: integer;
begin
  Assert(pDispIds <> nil);
  for i := 0 to dps.cArgs - 1 do
    pDispIds^[i] := dps.cArgs - 1 - i;
  if (dps.cNamedArgs <= 0) then Exit;
  for i := 0 to dps.cNamedArgs - 1 do
    pDispIds^[dps.rgdispidNamedArgs^[i]] := i;
end;

function  IsTopLevel(WB: IWebBrowser2): Boolean;
begin
    if not Assigned(WB) then
        Result := False
    else
    begin
        Result := WB.TopLevelContainer;
    end;
end;

function TwndMSAAV.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
type
  POleVariant = ^OleVariant;
var
  dps: TDispParams absolute Params;
  bHasParams: boolean;
  pDispIds: PDispIdList;
  iDispIdsSize: integer;
  i: Integer;
    TCP: PCP;
    pWB: IWebBrowser2;
    iDis: IDispatch;
    iDoc: IHTMLDocument2;
begin
  pDispIds := nil;
  iDispIdsSize := 0;
  bHasParams := (dps.cArgs > 0);
  if (bHasParams) then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end;
  try
    if (bHasParams) then BuildPositionalDispIds(pDispIds, dps);
    Result := S_OK;
    case DispId of
      {DISPID_STATUSTEXTCHANGE: DoStatusTextChange(dps.rgvarg^[pDispIds^[0]].bstrval);
      DISPID_PROGRESSCHANGE: DoProgressChange(dps.rgvarg^[pDispIds^[0]].lval, dps.rgvarg^[pDispIds^[1]].lval);
      DISPID_COMMANDSTATECHANGE: DoCommandStateChange(dps.rgvarg^[pDispIds^[0]].lval, dps.rgvarg^[pDispIds^[1]].vbool);
      DISPID_DOWNLOADBEGIN: DoDownloadBegin();
      DISPID_DOWNLOADCOMPLETE: DoDownloadComplete();
      DISPID_TITLECHANGE: DoTitleChange(dps.rgvarg^[pDispIds^[0]].bstrval);
      DISPID_PROPERTYCHANGE: DoPropertyChange(dps.rgvarg^[pDispIds^[0]].bstrval);}
      DISPID_BEFORENAVIGATE2:
      begin
        pWB := IDispatch(dps.rgvarg^[6].bstrVal) as IWebBrowser2;

        if IsTopLevel(pWB) then
        begin
            for i := 0 to FCPList.Count - 1 do
            begin
                if Assigned(FCPList.Items[i]) then
                begin
                    PCP(FCPList.Items[i])^.CP.Unadvise(PCP(FCPList.Items[i])^.Cookie);
                    Dispose(FCPList.Items[i]);
                end;
            end;
            FCPList.Clear;

        end;
    end;
      //DISPID_NEWWINDOW2: DoNewWindow2(IDispatch(dps.rgvarg^[pDispIds^[0]].pdispval^), dps.rgvarg^[pDispIds^[1]].pbool^);
     // DISPID_NAVIGATECOMPLETE2: DoNavigateComplete2(IDispatch(dps.rgvarg^[pDispIds^[0]].dispval), POleVariant(dps.rgvarg^[pDispIds^[1]].pvarval)^);
      DISPID_DOCUMENTCOMPLETE:
      begin
        pWB := IDispatch(dps.rgvarg^[1].bstrVal) as IWebBrowser2;
        if (POleVariant(dps.rgvarg^[pDispIds^[1]].pvarval)^ <> '') then
        begin
            if SUCCEEDED(pWB.Document.QueryInterface(IID_IHTMLDocument, iDis)) then
            begin
                if SUCCEEDED(iDis.QueryInterface(IID_IHTMLDocument2, iDoc)) then
                begin
                    if SUCCEEDED(iDis.QueryInterface(IID_IConnectionPointContainer, CPC)) then
                    begin
                        New(TCP);
                        if SUCCEEDED(CPC.FindConnectionPoint(DIID_HTMLDocumentEvents, TCP^.CP)) then
                        begin


                            TCP^.CP.Advise(Self, TCP^.Cookie);
                            FCPList.Add(TCP);
                        end
                        else
                            Dispose(TCP);

                    end;
                end;
            end;
        end;

      end;
      {DISPID_ONVISIBLE: DoOnVisible(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_ONTOOLBAR: DoOnToolBar(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_ONMENUBAR: DoOnMenuBar(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_ONSTATUSBAR: DoOnStatusBar(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_ONFULLSCREEN: DoOnFullScreen(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_ONTHEATERMODE: DoOnTheaterMode(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_WINDOWSETRESIZABLE: DoWindowSetResizable(dps.rgvarg^[pDispIds^[0]].vbool);
      DISPID_WINDOWCLOSING: DoWindowClosing(dps.rgvarg^[pDispIds^[0]].vbool, dps.rgvarg^[pDispIds^[1]].pbool^);
      DISPID_WINDOWSETLEFT: DoWindowSetLeft(dps.rgvarg^[pDispIds^[0]].lval);
      DISPID_WINDOWSETTOP: DoWindowSetTop(dps.rgvarg^[pDispIds^[0]].lval);
      DISPID_WINDOWSETWIDTH: DoWindowSetWidth(dps.rgvarg^[pDispIds^[0]].lval);
      DISPID_WINDOWSETHEIGHT: DoWindowSetHeight(dps.rgvarg^[pDispIds^[0]].lval);
      DISPID_CLIENTTOHOSTWINDOW: DoClientToHostWindow(dps.rgvarg^[pDispIds^[0]].plval^, dps.rgvarg^[pDispIds^[1]].plval^);
      DISPID_SETSECURELOCKICON: DoSetSecureLockIcon(dps.rgvarg^[pDispIds^[0]].lval);
      DISPID_FILEDOWNLOAD: DoFileDownload(dps.rgvarg^[pDispIds^[0]].pbool^); }
      DISPID_ONQUIT:
        begin
            //DoOnQuit();
            for i := 0 to FCPList.Count - 1 do
                begin
                    if Assigned(FCPList.Items[i]) then
                    begin
                        PCP(FCPList.Items[i])^.CP.Unadvise(PCP(FCPList.Items[i])^.Cookie);
                        Dispose(FCPList.Items[i]);
                    end;
                end;
                FCPList.Clear;
                CP.Unadvise(Cookie);
        end;
      {DISPID_HTMLELEMENTEVENTS2_ONFOCUSIN:
      begin
        OnElementFocus(self);
      end;
      //DISPID_HTMLELEMENTEVENTS_ONBLUR:
      DISPID_HTMLELEMENTEVENTS2_ONFOCUSOUT:
        OnElementBlur(self);  }
      //DISPID_HTMLELEMENTEVENTS2_ONMOUSEENTER:
      DISPID_HTMLELEMENTEVENTS2_ONMOUSEOVER:
      begin
        OnMouseEnter(self);
      end;
      //DISPID_HTMLELEMENTEVENTS2_ONMOUSELEAVE:
      DISPID_HTMLELEMENTEVENTS2_ONMOUSEOUT:
      begin
        OnMouseLeave(self);
      end;
    else
      Result := DISP_E_MEMBERNOTFOUND;
    end;
  finally
    if (bHasParams) then FreeMem(pDispIds, iDispIdsSize);
  end;
end;

procedure TwndMSAAV.OnMouseEnter(Sender: TObject);
var
    iWnd: IHTMLWindow2;
    iEvObj: IHTMLEventObj;
    iDoc: IHTMLDocument2;
begin
    //Beep;
    if not SUCCEEDED(IE.Document.QueryInterface(IID_IHTMLDocument2, iDoc)) then
        Exit;
    iWnd := IDoc.parentWindow;
    if iWnd = nil then
        Exit;
    if iWnd.event <> nil then
    begin
        iEvObj := iWnd.event;
    end;
    if iEvObj <> nil then
    begin
        iEVSrcEle := iEvObj.srcElement;
        //IE.StatusText := iEvObj.srcElement.tagName;
    end;
end;

procedure TwndMSAAV.OnMouseLeave(Sender: TObject);
begin
    iEVSrcEle := nil;
end;

procedure TwndMSAAV.PopupMenu2Popup(Sender: TObject);
begin

  if TreeView1.Items.Count = 0 then
  begin
  	mnuSAll.Enabled := false;
    mnuSSel.Enabled := false
  end
  else
  begin
    mnuSAll.Enabled := True;
    if TreeView1.SelectionCount = 0 then
  		mnuSSel.Enabled := false
  	else
    	mnuSSel.Enabled := True;
  end;
end;

procedure TwndMSAAV.ThDone(Sender: TObject);
begin
    if not bTer then
    begin

        if (not acMSAAMode.Checked) then
        begin
            if iMode <> 2 then
            begin

                if (acShowTip.Checked) then
                begin
                    ShowTipWnd;
                end;
            end;
        end
        else
        begin

            if (acShowTip.Checked) then
            begin
                ShowTipWnd;
            end;
        end;
    end;

end;
end.
