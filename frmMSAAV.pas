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
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, System.UITypes, System.Types,
  Dialogs, ImgList, ComCtrls, ToolWin,  StdCtrls, ShellAPI,
  IniFiles, ActnList, CommCtrl, FocusRectWnd, ClipBrd,
  ActiveX, MSHTML_tlb, ExtCtrls, Shlobj, WinAPI.oleacc,
  iAccessible2Lib_tlb, ISimpleDOM, Actions,
  Menus, Thread, TipWnd, UIAutomationClient_TLB, AccCTRLs,
  frmSet, Math, IntList, StrUtils, PermonitorApi, Multimon, VirtualTrees,
  System.ImageList, UIA_TLB, SynEdit, SynEditHighlighter, SynHighlighterHtml, CommDlg;

const
    IID_IServiceProvider: TGUID = '{6D5140C1-7436-11CE-8034-00AA006009FA}';
   	IID_IStylesProvider: TGUID = '{19b6b649-f5d7-4a6d-bdcb-129252be588a}';
    P2WDEF = 350;
    P4HDEF = 250;



type
	IStylesProvider = interface;
  IStylesProvider = interface(IUnknown)
  ['{19b6b649-f5d7-4a6d-bdcb-129252be588a}']
   	function get_StyleId(out retVal: SYSINT): HResult; stdcall;
		function get_StyleName(out retVal: WideString): HResult; stdcall;
    function get_FillColor(out retVal: SYSINT): HResult; stdcall;
    function get_FillPatternStyle(out retVal: WideString): HResult; stdcall;
    function get_Shape(out retVal: WideString): HResult; stdcall;
    function get_FillPatternColor(out retVal: SYSINT): HResult; stdcall;
    function get_ExtendedProperties(out retVal: WideString): HResult; stdcall;
  end;
  PSCData = ^TSCData;
  TSCData = record
  	Name: String;
    SCKey: TShortCut;
    actName: String;
  end;
   PNodeData = ^TNodeData;
   TNodeData = record
    Value1: String;
    Value2: string;
    Acc: IAccessible;
    iID: integer;

   end;
   PTreeData = ^TTreeData;
   TTreeData = record
      Acc: IAccessible;
      uiEle: IUIAUTOMATIONELEMENT;
      iID: integer;
      dummy: boolean;
   end;
   {PIE = ^FIE;
    FIE = record
        Iweb: IWebBrowser2;
    end; }

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

  TwndMSAAV = class(TForm, IUIAutomationFocusChangedEventHandler)
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
    PB1: TAccClrBtn;
    Panel4: TPanel;
    PB3: TAccClrBtn;
    Panel5: TPanel;
    Splitter2: TSplitter;
    Panel3: TPanel;
    PB2: TAccClrBtn;
    acTVcol: TAction;
    acTLCol: TAction;
    acMMCol: TAction;
    mnuColl: TMenuItem;
    mnuTV: TMenuItem;
    mnuTL: TMenuItem;
    mnuCode: TMenuItem;
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
    PopupMenu2: TPopupMenu;
    mnuSAll: TMenuItem;
    mnuSSel: TMenuItem;
    mnuTVSAll: TMenuItem;
    mnuTVSSel: TMenuItem;
    mnuSave: TMenuItem;
    mnuOpenB: TMenuItem;
    mnuOAll: TMenuItem;
    mnuOSel: TMenuItem;
    mnuTVOpen: TMenuItem;
    mnuTVOAll: TMenuItem;
    mnuTVOSel: TMenuItem;
    mnuSelMode: TMenuItem;
    TreeList1: TVirtualStringTree;
    acTreeFocus: TAction;
    acListFocus: TAction;
    acMemoFocus: TAction;
    ac3ctrls: TAction;
    ImageList4: TImageList;
    Memo1: TAccSynEdit;
    SynHTMLSyn1: TSynHTMLSyn;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    tbUIA: TAccTreeView;
    PopupMenu3: TPopupMenu;
    mnutvMSAA: TMenuItem;
    mnutvUIA: TMenuItem;
    mnutvBoth: TMenuItem;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure acFocusExecute(Sender: TObject);
    procedure acCursorExecute(Sender: TObject);
    procedure acCopyExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure acRectExecute(Sender: TObject);
    procedure acOnlyFocusExecute(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
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
    procedure mnuTVOAllClick(Sender: TObject);
    procedure mnuOAllClick(Sender: TObject);
    procedure mnuTVOSelClick(Sender: TObject);
    procedure mnuOSelClick(Sender: TObject);
    procedure mnuSelModeClick(Sender: TObject);
    procedure TreeList1GetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure acTreeFocusExecute(Sender: TObject);
    procedure acListFocusExecute(Sender: TObject);
    procedure acMemoFocusExecute(Sender: TObject);
    procedure ac3ctrlsExecute(Sender: TObject);
    procedure TreeList1FreeNode(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure tbUIADeletion(Sender: TObject; Node: TTreeNode);
    procedure tbUIAChange(Sender: TObject; Node: TTreeNode);
    procedure tbUIAAddition(Sender: TObject; Node: TTreeNode);
    procedure PageControl1Change(Sender: TObject);
    procedure mnutvUIAClick(Sender: TObject);
  private
    { Private declarations }

    HTMLsFF: array [0..2, 0..1] of string;
    ARIAs: array [0..1, 0..1] of string;
    IA2Sts: array [0..17] of string;
    QSFailed: string;
    sHTML, sTxt, sTypeIE, sTypeFF, sARIA, NodeTxt, Err_Inter: string;
    bSelMode: Boolean;
    iFocus, iRefCnt, ShowSrcLen, iPID: integer;
    hHook: THandle;
    LangList, ClsNames: TStringList;
    rType, rTarg: string;
    arPT: array [0..2] of TPoint;

    Roles: array [0..42] of string;
    StyleID: array [0..16] of string;
    oldPT: TPoint;
    WndFocus, WndLabel, WndDesc, WndTarg: TwndFocusRect;
    WndTip: TfrmTipWnd;
    hRgn1, hRgn2, hRgn3: hRgn;

    Created: boolean;
    CEle: IHTMLElement;
    SDom: ISImpleDOMNode;
    UIEle: IUIAutomationElement;
    TreeTH: TreeThread;
    UIATH: TreeThread4UIA;
    HTMLTh: HTMLThread;

    Treemode, bPFunc: boolean;
    //UIA: IUIAutomation;
    P2W, P4H: integer;
    cDPI, iTHCnt: integer;
    bFirstTime, bTabEvt: boolean;
    TipText, TipTextIA2, sFilter: string;
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
    ExTip, nSelected: boolean;
    DefFont, DefW, DefH: integer;
    ScaleX, ScaleY, DefX, DefY: double;
    sMSAAtxt, sHTMLtxt, sUIATxt, sARIATxt, sIA2Txt: string;
    dEventTime: dword;
    function SaveHTMLDLG(initDir: string; var FName: string): boolean;
    procedure ExecOnlyFocus;
    function MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = false): string;
    function MSAAText4UIA(TextOnly: boolean = false): string;
    function MSAAText4Tip(pAcc: IAccessible = nil): string;
    function MSAAText4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
    function IA2Text4HTML(pAcc: IAccessible = nil; tab: string = ''): string;
    function HTMLText: string;
    function HTMLText4FF: string;
    function ARIAText: string;
    function SetNodeData(pNode: PVirtualNode; Text1, Text2: string; Acc: iAccessible; iID: integer): PVirtualNode;
    function SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
    function UIAText(HTMLout: boolean = false; tab: string = ''): string;
    function GetEleName(Acc: IAccessible): string;
    procedure SizeChange;
    procedure ExecMSAAMode(sender: TObject);
    //Thread terminate
    procedure ThDone(Sender: TObject);
    procedure ThHTMLDone(Sender: TObject);
    procedure RecursiveID(sID: string; isEle: ISimpleDOMNODE; Labelled: boolean = true);
    procedure ShowBalloonTip(Control: TWinControl; Icon: integer; Title: string; Text: string; RC: TRect; X, Y:integer; Track:boolean = false);
    procedure SetBalloonPos(X, Y:integer);
    procedure RecursiveACC(ParentNode: TTreeNode; ParentAcc: IAccessible);
    procedure RecursiveIA2Rel(pNode: pVirtualNode; ia2: IAccessible2);
    function GetSameAcc: boolean;
    function IsSameUIElement(ia1, ia2: IAccessible; iID1, iID2: integer): boolean;
    procedure SetTreeMode(pNode: TTreeNode);
    procedure mnuLangChildClick(Sender: TObject);
    procedure RecursiveTV(cNode: TTreeNode; var HTML: string; var iCnt: integer; ForSel: boolean = false);
    function Get_RoleText(Acc: IAccessible; Child: integer): string;
    function GetIA2Role(iRole: integer): string;
    function GetIA2State(iRole: integer): string;
    procedure ExecCmdLine;
    function GetTVAllItems: string;
    function GetTVSelItems: string;
    procedure WMDPIChanged(var Message: TMessage); message WM_DPICHANGED;
  protected
    { Protected declarations  }

  public
    { Public declarations }
    ConvErr, HelpURL, DllPath, APPDir, sTrue, sFalse: string;
    procedure GetNaviState(AllFalse: boolean = false);
    procedure Load;
    function LoadLang: string;
    procedure ShowRectWnd(bkClr: TColor);
    procedure ShowRectWnd2(bkClr: TColor; RC:TRect);
    procedure ShowLabeledWnd(bkClr: TColor; RC:TRect);
    procedure ShowDescWnd(bkClr: TColor; RC:TRect);
    procedure ShowTargWnd(bkClr: TColor; RC:TRect);
    procedure ShowTipWnd;
    function ShowMSAAText:boolean;
    procedure ShowText4UIA;
    procedure WMCopyData(var Msg: TWMCopyData); Message WM_COPYDATA;
    Procedure SetAbsoluteForegroundWindow(HWND: hWnd);
    function GetWindowNameLC(Wnd: HWND): string;

    function HandleFocusChangedEvent(const sender: IUIAutomationElement): HResult; stdcall;
  end;

var
  wndMSAAV: TwndMSAAV;
  ac: IAccessible;
  DMode: boolean;
  bTer: boolean;
  refNode, LoopNode, sNode: TTreeNode;
  refVNode: PVirtualNode;
  TBList, uTBList: TIntegerList;
  lMSAA: array[0..11] of string;
  lIA2: array [0..15] of string;
  lUIA: array [0..61] of string;
  HTMLs: array [0..2, 0..1] of string;
  flgMSAA, flgIA2, flgUIA, flgUIA2: integer;
  scList: TList;
  None: string;
implementation


{$R *.dfm}

function DoubleToInt(d: double): integer;
begin
  SetRoundMode(rmUP);
  Result := Trunc(SimpleRoundTo(d));
end;

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
  Result := '';
  if SUCCEEDED(SHGetSpecialFolderLocation(Application.Handle, CSIDL_PERSONAL, IIDList)) then
  begin
    if not SHGetPathFromIDList(IIDList, buffer) then
    begin
      raise Exception.Create('A virtual diractory cannot be acquired.');
    end
    else
      Result := StrPas(Buffer);
  end;
end;

procedure WinEventProc(hWinEventHook: THandle; event: DWORD; HWND: HWND;
	idObject, idChild: Longint; idEventThread, dwmsEventTime: DWORD); stdcall;
var

	vChild: variant;
	pAcc: IAccessible;
	iDis: iDispatch;
  hr: HResult;


begin
	if wndMSAAV.bPFunc then Exit;

	hr := AccessibleObjectFromEvent(HWND, idObject, idChild, @pAcc, vChild);
  if (hr = 0) and (Assigned(pAcc)) and (event = EVENT_OBJECT_FOCUS) and (wndMSAAV.dEventTime < dwmsEventTime) then
  begin
  	wndMSAAV.UIAuto.ElementFromiAccessible(pAcc, vChild, wndMSAAV.uiEle);
  	try
			wndMSAAV.dEventTime := dwmsEventTime;

			if (wndMSAAV.acOnlyFocus.Checked) then
			begin
				wndMSAAV.VarParent := vChild;
				wndMSAAV.iAcc := pAcc;
				pAcc.Get_accParent(iDis);
				wndMSAAV.accRoot := iDis as IAccessible;
				wndMSAAV.ShowRectWnd(clYellow);
			end
			else if (wndMSAAV.acFocus.Checked) then
			begin
				wndMSAAV.VarParent := vChild;
				wndMSAAV.iAcc := pAcc;
				pAcc.Get_accParent(iDis);
				wndMSAAV.accRoot := iDis as IAccessible;
				wndMSAAV.Timer2.Enabled := false;
				wndMSAAV.Timer2.Enabled := True;
			end;
		except
			on E: Exception do
				ShowErr(E.Message);
		end;
  end;

end;

function TwndMSAAV.HandleFocusChangedEvent(const sender: IUIAutomationElement): HResult;
var
    hr: HResult;
    icPID: integer;
    rc: UIAutomationClient_tlb.tagRECT;

begin
  Result := S_OK;
  if bPFunc then Exit;

  UIEle := sender;
  if (Assigned(uiEle)) then
  begin
  	hr := uiEle.Get_CurrentProcessId(icPID);
    if (hr = S_OK) and (icPID = wndMSAAV.iPID) then
    	exit;
  end;

  if (mnutvUIA.Checked) then
  begin
  	if (wndMSAAV.acOnlyFocus.Checked) then
		begin
			uiEle.Get_CurrentBoundingRectangle(RC);
			ShowRectWnd2(clYellow, Rect(RC.left, RC.top, RC.right - RC.left,
				RC.bottom - RC.top));
		end
		else if (wndMSAAV.acFocus.Checked) then
		begin
      if wndMSAAV.acRect.Checked then
      begin
				uiEle.Get_CurrentBoundingRectangle(RC);
				ShowRectWnd2(clRed, Rect(RC.left, RC.top, RC.right - RC.left,
					RC.bottom - RC.top));
      end;
			Timer2.Enabled := false;
			Timer2.Enabled := True;
		end;
  end;
end;


procedure TwndMSAAV.Timer1Timer(Sender: TObject);
var
	Wnd: HWND;
	pAcc: IAccessible;
	v: variant;
	iDis: iDispatch;
	hr: HResult;
  icPID: integer;
  tagPT: UIAutomationClient_TLB.tagPoint;
begin
	if (not acCursor.Checked) or bPFunc then
		Exit;


	arPT[2] := arPT[1];
	arPT[1] := arPT[0];
	getcursorpos(arPT[0]);

	if (arPT[0].X = oldPT.X) and (arPT[0].Y = oldPT.Y) then
		Exit;
	try
		if (arPT[0].X = arPT[1].X) and (arPT[0].X = arPT[2].X) and
			(arPT[2].X = arPT[1].X) and (arPT[0].Y = arPT[1].Y) and
			(arPT[0].Y = arPT[2].Y) and (arPT[2].Y = arPT[1].Y) then
		begin
      hr := AccessibleObjectFrompoint(arPT[0], @pAcc, v);
			if (hr = 0) and (Assigned(pAcc)) then
			begin
				if ((not acOnlyFocus.Checked) and (acCursor.Checked)) then
				begin
					if SUCCEEDED(WindowFromAccessibleObject(pAcc, Wnd)) then
					begin


						//hr := UIAuto.ElementFromiAccessible(pAcc, v, uiEle);
            tagPT.X := arPT[0].X;
						tagPT.Y := arPT[0].Y;
						hr := UIAuto.ElementFromPoint(tagPT, uiEle);
            if (hr = 0) and (Assigned(uiEle)) then
            begin
              hr := uiEle.Get_CurrentProcessId(icPID);
              if (hr = S_OK) and (icPID = iPID) then
              	exit;
            end;
						Treemode := false;
						iAcc := pAcc;
						VarParent := v;

						if iAcc <> nil then
						begin
							oldPT := arPT[0];
							arPT[0] := Point(0, 0);
							arPT[1] := Point(0, 0);
							arPT[2] := Point(0, 0);
						end;

						hr := pAcc.Get_accParent(iDis);
						if (hr = 0) and (Assigned(iDis)) then
						begin
							hr := iDis.QueryInterface(IID_IACCESSIBLE, accRoot);
							if (hr = 0) and (Assigned(accRoot)) then
								accRoot := iDis as IAccessible;
						end;
            iTHCnt := 0;
            bTabEvt := false;
            if not mnutvUIA.Checked then
            begin
							PageControl1.ActivePageIndex := 0;
              TabSheet1.TabVisible := True;
              TabSheet2.TabVisible := False;
              acShowTip.Enabled := True;
							ShowMSAAText;
            end
            else
            begin
              PageControl1.ActivePageIndex := 1;
              TabSheet1.TabVisible := False;
              TabSheet2.TabVisible := True;
              acShowTip.Enabled := False;
            	ShowText4UIA;
            end;
          end
          else
          	iAcc := nil;
				end;
			end;
		end;
	except

	end;

end;


procedure TwndMSAAV.Timer2Timer(Sender: TObject);
begin
	Timer2.Enabled := false;
	Treemode := false;

	if (not acOnlyFocus.Checked) and (acFocus.Checked) then
	begin
  	iTHCnt := 0;
    bTabEvt := True;
		if not mnutvUIA.Checked then
		begin
			PageControl1.ActivePageIndex := 0;
			TabSheet1.TabVisible := True;
      TabSheet2.TabVisible := False;
			acShowTip.Enabled := True;
			ShowMSAAText;
		end
		else
		begin
			PageControl1.ActivePageIndex := 1;
			TabSheet1.TabVisible := False;
      TabSheet2.TabVisible := True;
			acShowTip.Enabled := false;
			ShowText4UIA;
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
  PC:PChar;
  ovValue: OleVariant;

begin

  case iRole of
    1025..1067: result := Roles[iRole - 1025];
    else
    begin
      PC := StrAlloc(255);
      ovValue := iROle;
      GetRoleTextW(ovValue, PC, StrBufSize(PC));
      result := PC;
      StrDispose(PC);
    end;
  end;
end;

function TwndMSAAV.GetIA2State(iRole: integer): string;
begin
  Result := '';

            if (iRole and IA2_STATE_ACTIVE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[0];
            end;

            if (iRole and IA2_STATE_ARMED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[1];
            end;

            if (iRole and IA2_STATE_DEFUNCT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[2];
            end;

            if (iRole and IA2_STATE_EDITABLE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[3];
            end;

            if (iRole and IA2_STATE_HORIZONTAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[4];
            end;

            if (iRole and IA2_STATE_ICONIFIED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[5];
            end;

            if (iRole and IA2_STATE_INVALID_ENTRY) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[6];
            end;

            if (iRole and IA2_STATE_MANAGES_DESCENDANTS) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[7];
            end;

            if (iRole and IA2_STATE_MODAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[8];
            end;

            if (iRole and IA2_STATE_MULTI_LINE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[9];
            end;

            if (iRole and IA2_STATE_OPAQUE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[10];
            end;

            if (iRole and IA2_STATE_REQUIRED) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[11];
            end;

            if (iRole and IA2_STATE_SELECTABLE_TEXT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[12];
            end;

            if (iRole and IA2_STATE_SINGLE_LINE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[13];
            end;

            if (iRole and IA2_STATE_STALE) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[14];
            end;

            if (iRole and IA2_STATE_SUPPORTS_AUTOCOMPLETION) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[15];
            end;

            if (iRole and IA2_STATE_TRANSIENT) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[16];
            end;

            if (iRole and IA2_STATE_VERTICAL) <> 0 then
            begin
                if Result <> '' then Result := result + ', ';
                Result := Result + IA2Sts[17];
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

procedure TwndMSAAV.RecursiveACC(ParentNode: TTreeNode; ParentAcc: IAccessible);
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
        refNode := TreeView1.Items.AddChildObject(pNode, NodeCap, Pointer(TD));
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

            sNode := refNode;
        end;

        cNode := refNode;

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
                RecursiveACC(cNode, iDis as IAccessible);
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

procedure TwndMSAAV.ShowText4UIA;
var
  pEle: IUIAutomationElement;
  bSame: boolean;
  hr: hresult;
	iLeg: IUIAutomationLegacyIAccessiblePattern;
	iInt: IInterface;
	iSP: IServiceProvider;
  tAcc: IAccessible;
  iEle: IHTMLElement;
  iDom: IHTMLDomNode;
  iSEle: ISimpleDOMNode;
  eleHWND, paHWND: HWND;
  gtcStart, gtcStop: cardinal;
	function GetiLegacy: HResult;
	begin
  	result := S_FALSE;
  	hr := uiEle.GetCurrentPattern(UIA_LegacyIAccessiblePatternId, iInt);
		if (hr = 0) and (Assigned(iInt)) then
		begin
    	result := iInt.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern, iLeg);
		end;

	end;
  procedure GetSameUIEle;
  var
    i, iSame: integer;
    rNode: TTreeNode;
  begin
    bSame := false;
    try
        if uTBList.Count > 0 then
        begin
          for i := 0 to uTBList.Count - 1 do
          begin
            rNode := tbUIA.Items.GetNode(HTreeItem(uTBList.Items[i]));
            if (Assigned(rNode)) and (SUCCEEDED(UIAuto.CompareElements(uiEle, TTreeData(rNode.Data^).uiEle, iSame))) then
            begin
              if iSame <> 0 then
              begin
                bSame := True;
                tbUIA.OnChange := nil;
                tbUIA.SetFocus;
                tbUIA.TopItem := rNode;
                rNode.Expanded := True;
                rNode.Selected := True;
                // GetNaviState;
                if (acShowTip.Checked) then
                begin
                  ShowTipWnd;
                end;
                Break;
              end;
            end;
            Application.ProcessMessages;
          end;
        end;
    finally
      tbUIA.OnChange := tbUIAChange;
    end;
  end;
begin

  if (mnutvUIA.Checked) then
  begin
  	TreeList1.BeginUpdate;
		TreeList1.Clear;
		TreeList1.EndUpdate;
		mnuTVSAll.Enabled := false;
		mnuTVOAll.Enabled := false;
		mnuTVSSel.Enabled := false;
		mnuTVOSel.Enabled := false;
    Memo1.ClearAll;
  end
  else
  begin
  	if TreeMode then
    begin
    	TreeList1.BeginUpdate;
		TreeList1.Clear;
		TreeList1.EndUpdate;
    Memo1.ClearAll;
    end;
  end;

  iDefIndex := -1;

  bSame := False;
  pEle := nil;
  try

      if Assigned(UIATH) then
      begin
        UIATH.Terminate;
        UIATH.WaitFor;
        UIATH.Free;
        UIATH := nil;
      end;
      if (mnutvUIA.Checked) or (TreeMode) then
      	UIAText;

      hr := GetiLegacy;
      if (hr = 0) and (Assigned(iLeg)) then
      begin
      	hr := iLeg.GetIAccessible(tAcc);
        if (hr = 0) and (Assigned(tAcc))then
        begin
          if (mnuAll.Checked) and (SUCCEEDED(WindowFromAccessibleObject(tAcc, eleHWND))) then
          begin
          	paHWND := GetAncestor(eleHWND, GA_ROOT);
            UIAuto.ElementFromHandle(pointer(paHWND), pEle);
          end;
          if (mnuMSAA.Checked) and (mnutvUIA.Checked) then
          	sMSAATxt := MSAAText(tAcc);
          hr := iAcc.QueryInterface(IID_IServiceProvider, iSP);
					if SUCCEEDED(hr) and Assigned(iSP) then
					begin
						hr := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
						if SUCCEEDED(hr) and Assigned(iEle) then
						begin
							hr := iEle.QueryInterface(IID_IHTMLDOMNode, iDom);
							if SUCCEEDED(hr) and Assigned(iDom) then
							begin
								CEle := iEle;
                if (mnutvUIA.Checked)  or (TreeMode) then
								begin
									if mnuARIA.Checked then
									begin
										ARIAText;
									end;
									if mnuHTML.Checked then
									begin

										HTMLText;
									end;
								end;
              end;
            end
            else
						begin
							hr := iSP.QueryService(IID_ISIMPLEDOMNODE,
								IID_ISIMPLEDOMNODE, isEle);
							if SUCCEEDED(hr) and Assigned(isEle) then
							begin
								SDom := isEle;
								// GetNaviState;
                if (mnutvUIA.Checked) or (TreeMode) then
								begin
									if mnuARIA.Checked then
										ARIAText;
									if mnuHTML.Checked then
										HTMLText4FF;
								end;
							end;

						end;
          end;
          if (mnutvUIA.Checked) or (TreeMode) then
          begin
          	if (mnuIA2.Checked) then
							SetIA2Text;
          end;
        end
        else
        begin
        	if (mnutvUIA.Checked) or (TreeMode) then
          begin
        		if (mnuMSAA.Checked) then
        			sMSAATxt := MSAAText4UIA;
          end;
        end;
      end;
      //if mnuMSAA.Checked then
      //sMSAATxt := MSAAText4UIA;
      //if mnuUIA.Checked then


      if (not Treemode) and (Splitter1.Enabled = True) then
      begin
        if (tbUIA.Items.Count > 0) then
        begin
          Timer1.Enabled := false;
          try
            GetSameUIEle;
          finally
            Timer1.Enabled := True;
          end;
        end;

        if not bSame then
        begin
          uTBList.Clear;
          tbUIA.Items.Clear;
          if bTabEvt then
					begin
						gtcStart := GetTickCount;
						gtcStop := GetTickCount;
						while ((gtcStop - gtcStart) < 2000) do
						begin
							Application.ProcessMessages;
							gtcStop := GetTickCount;
						end;
					end;
          if not Assigned(pEle) then
            pEle := uiEle;
          if mnuAll.Checked then
          	UIATH := TreeThread4UIA.Create(UIAuto, uiEle, pEle, None, True, true, nil)
          else
          	UIATH := TreeThread4UIA.Create(UIAuto, uiEle, pEle, None, True, False, nil);
          UIATH.OnTerminate := ThDone;
          UIATH.Start;
        end;
      end;
      if mnutvUIA.Checked then
			begin
				if TreeList1.ChildCount[TreeList1.RootNode] > 0 then
				begin
					TreeList1.Expanded[TreeList1.RootNode];
					TreeList1.OffsetX := 0;
					TreeList1.OffsetY := 0;
					TreeList1.Selected[TreeList1.RootNode] := True;

				end;
				if Treemode then
					GetNaviState;
			end;
  except
    on E: Exception do
      ShowErr(E.Message);
  end;
end;


function TwndMSAAV.ShowMSAAText:boolean;
var
    Wnd: hwnd;
    iSP: IServiceProvider;
    iEle, paEle: IHTMLElement;
    Path, cAccRole: string;
    ws: widestring;
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
    gtcStart, gtcStop: cardinal;
    procedure GetAccTxt;
		begin
			iAcc.Get_accName(VarParent, ws);
      NodeTxt := String(ws);
			if NodeTxt = '' then
				NodeTxt := None;
			NodeTxt := StringReplace(NodeTxt, #13, ' ', [rfReplaceAll]);
			NodeTxt := StringReplace(NodeTxt, #10, ' ', [rfReplaceAll]);
			cAccRole := Get_RoleText(iAcc, VarParent);
			NodeTxt := NodeTxt + ' - ' + cAccRole;

		end;
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
    Result := false;

	if not Assigned(iAcc) then
		Exit;
  if (not mnutvUIA.Checked) then
  begin
		TreeList1.BeginUpdate;
		TreeList1.Clear;
		TreeList1.EndUpdate;
		mnuTVSAll.Enabled := false;
		mnuTVOAll.Enabled := false;
		mnuTVSSel.Enabled := false;
		mnuTVOSel.Enabled := false;
    Memo1.ClearAll;
  end
  else
  begin
  	if TreeMode then
    begin
    	TreeList1.BeginUpdate;
		TreeList1.Clear;
		TreeList1.EndUpdate;
    Memo1.ClearAll;
    end;
  end;
	iDefIndex := -1;
	LoopNode := nil;

  sMSAAtxt := '';
  sHTMLtxt := '';
  sUIATxt := '';
  sARIATxt := '';
  sIA2Txt := '';
  nSelected := False;


		GetNaviState(True);
		CEle := nil;
		SDom := nil;
		Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
		if acOnlyFocus.Checked then
			ShowRectWnd(clYellow)
		else if (acRect.Checked) then
		begin
			ShowRectWnd(clRed);
		end;
		if Assigned(TreeTH) then
		begin
			TreeTH.Terminate;
			TreeTH.WaitFor;
			TreeTH.Free;
			TreeTH := nil;
		end;


			if (mnuMSAA.Checked) and ((not mnutvUIA.Checked) or (TreeMode)) then
				sMSAAtxt := MSAAText;

			iRes := iAcc.QueryInterface(IID_IServiceProvider, iSP);
			if SUCCEEDED(iRes) and Assigned(iSP) then
			begin
				iRes := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
				if SUCCEEDED(iRes) and Assigned(iEle) then
				begin

						CEle := iEle;
            if (not mnutvUIA.Checked) or (TreeMode) then
						begin
							if mnuARIA.Checked then
							begin
								sARIATxt := ARIAText;
							end;
							if mnuHTML.Checked then
							begin
								HTMLText;
							end;
						end;

						if (not Treemode) and (Splitter1.Enabled = True) then
						begin
							iDoc := iEle.Document as IHTMLDocument2;

							paEle := iEle;
							if mnuAll.Checked then
							begin
								paEle := iDoc.body;

							end;
							if Assigned(paEle) then
							begin
								iRes := paEle.QueryInterface(IID_IServiceProvider, iSP);
								if SUCCEEDED(iRes) and Assigned(iSP) then
								begin
									iRes := iSP.QueryService(IID_IACCESSIBLE,
										IID_IACCESSIBLE, pAcc);
									if SUCCEEDED(iRes) and Assigned(pAcc) then
									begin

										bSame := false;
										if (mnuAll.Checked) and (IsSameUIElement(DocAcc, pAcc, 0, 0))
										then
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
											iAcc.accLocation(sRC.left, sRC.top, sRC.right,
												sRC.bottom, 0);
                      GetAccTxt;
											WindowFromAccessibleObject(iAcc, AWnd);
											TBList.Clear;

                      if bTabEvt then
                      begin
                      	gtcStart := GetTickCount;
                        gtcStop := GetTickCount;
                        while ((gtcStop - gtcStart) < 2000) do
                        begin
                        	application.ProcessMessages;
                          gtcStop := GetTickCount;
                        end;
                      end;
											TreeView1.Items.Clear;
											TreeTH := TreeThread.Create(iAcc, pAcc, None,
												True, mnuAll.Checked, nil, VarParent);
											TreeTH.OnTerminate := ThDone;
											TreeTH.Start;

										end;
									end;
								end;

							end;
						end;
				end
				else
				begin
					iRes := iSP.QueryService(IID_ISIMPLEDOMNODE,
						IID_ISIMPLEDOMNODE, isEle);
					if SUCCEEDED(iRes) and Assigned(isEle) then
					begin
						SDom := isEle;
						// GetNaviState;
            if (not mnutvUIA.Checked) or (TreeMode) then
						begin
							if mnuARIA.Checked then
								sARIATxt := ARIAText;
							if mnuHTML.Checked then
								sHTMLtxt := HTMLText4FF;
						end;
						if (not Treemode) and (Splitter1.Enabled = True) then
						begin
							iRes := iSP.QueryService(IID_ISIMPLEDOMNODE,
								IID_ISIMPLEDOMNODE, isEle);
							if SUCCEEDED(iRes) and Assigned(isEle) then
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
										end;
										inc(i);
										if i > 500 then
										begin
											isEle := nil;
											break;
										end;
										// ROLE_SYSTEM_DOCUMENT
									end; // end while
								end;
								if Assigned(isEle) then
								begin
									iRes := isEle.QueryInterface(IID_IServiceProvider, iSP);
									if SUCCEEDED(iRes) and Assigned(iSP) then
									begin
										iRes := iSP.QueryService(IID_IACCESSIBLE,
											IID_IACCESSIBLE, pAcc);
										if SUCCEEDED(iRes) and Assigned(pAcc) then
										begin
											// pAcc := iAcc;
											bSame := false;
											if (mnuAll.Checked) and
												(IsSameUIElement(DocAcc, pAcc, 0, 0)) then
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
												iAcc.accLocation(sRC.left, sRC.top, sRC.right,
													sRC.bottom, 0);
                        if not mnutvUIA.Checked then
                        	GetAccTxt;
												WindowFromAccessibleObject(iAcc, AWnd);
												TreeView1.Items.Clear;
                        if bTabEvt then
												begin
													gtcStart := GetTickCount;
													gtcStop := GetTickCount;
													while ((gtcStop - gtcStart) < 2000) do
													begin
														Application.ProcessMessages;
														gtcStop := GetTickCount;
													end;
												end;
												TreeTH := TreeThread.Create(iAcc, pAcc,
													None, True, mnuAll.Checked, nil, 0);
												TreeTH.OnTerminate := ThDone;
												TreeTH.Start;
											end;
										end;
									end;
								end; // isEle <> nil End
							end; // if SUCCEEDED(iRes) and Assigned(isEle) then End

						end; // IaEle
					end
					else
					begin
						Timer1.Enabled := false;
						try
							bSame := GetSameAcc;
						finally
							Timer1.Enabled := True;
						end;
						if not bSame then
						begin
							iAcc.accLocation(sRC.left, sRC.top, sRC.right, sRC.bottom, 0);
							WindowFromAccessibleObject(iAcc, AWnd);
							TBList.Clear;
							TreeView1.Items.Clear;
              if bTabEvt then
							begin
								gtcStart := GetTickCount;
								gtcStop := GetTickCount;
								while ((gtcStop - gtcStart) < 2000) do
								begin
									Application.ProcessMessages;
									gtcStop := GetTickCount;
								end;
							end;
							TreeTH := TreeThread.Create(iAcc, iAcc, None, True,
								false, nil, VarParent);
							TreeTH.OnTerminate := ThDone;
							TreeTH.Start;
						end;
					end;
				end;
			end;
      if (not mnutvUIA.Checked) or (TreeMode) then
  		begin
				if mnuIA2.Checked then
					sIA2Txt := SetIA2Text;
				if mnuUIA.Checked then
					sUIATxt := UIAText;
				if Treemode then
					GetNaviState;
        if TreeList1.ChildCount[TreeList1.RootNode] > 0 then
				begin
					TreeList1.Expanded[TreeList1.RootNode];
					TreeList1.OffsetX := 0;
					TreeList1.OffsetY := 0;
					TreeList1.Selected[TreeList1.RootNode] := True;
					// GetNaviState;
				end;
      end;

		Result := True;
end;



procedure TwndMSAAV.RecursiveIA2Rel(pNode: PVirtualnode; ia2: IAccessible2);
var
    isp: iserviceprovider;
    iInter: IInterface;
    ia2Targ: iaccessible2;
    iaTarg: IAccessible;
    hr, iRole, t, p, iTarg, iRel: integer;
    cNode, ccNode, cccNode: PVirtualNode;
    iAL: IAccessibleRelation;
    s, ws: widestring;
    ND: PNodeData;
begin
    Inc(iRefCnt);
    if iRefCnt > 5 then
        exit;
    try
        if SUCCEEDED(ia2.Get_nRelations(iRole)) then
        begin
            for p := 0 to iRole - 1 do
            begin

                ia2.Get_relation(p, iAL);
                //cnode := nodes.AddChild(pNode, inttostr(p));
                cnode := TreeList1.InsertNode(pNode, amAddChildLast, nil);
                ND := TreeList1.GetNodeData(cnode);
                ND.Value1 := inttostr(p);
                ND.Value2 := '';
                ND.Acc := nil;

                iAL.Get_RelationType(s);

                {ccNode := nodes.AddChild(cNode, rType);
                TreeList1.SetNodeColumn(ccnode, 1, s);
                ccNode := nodes.AddChild(cNode, rTarg);}
                ccnode := TreeList1.InsertNode(cNode, amAddChildLast, nil);
                ND := TreeList1.GetNodeData(ccnode);
                ND.Value1 := rType;
                ND.Value2 := s;
                ND.Acc := nil;

                ccnode := TreeList1.InsertNode(cNode, amAddChildLast, nil);
                ND := TreeList1.GetNodeData(ccnode);
                ND.Value1 := rTarg;
                ND.Value2 := '';
                ND.Acc := nil;

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
                                    {cccNode := nodes.AddChild(ccNode, inttostr(t));
                                    TreeList1.SetNodeColumn(cccNode, 1, s + ' - ' + ws);
                                    New(TD);
                                    TD^.Acc := iaTarg;
                                    TD^.iID := 0;
                                    cccNode.Data := Pointer(TD); }

                                    cccnode := TreeList1.InsertNode(ccNode, amAddChildLast, nil);
                                    ND := TreeList1.GetNodeData(cccnode);
                                    ND.Value1 := inttostr(t);
                                    ND.Value2 := s + ' - ' + ws;
                                    ND.Acc := iaTarg;
                                    ND.iID := 0;
                                    if SUCCEEDED(ia2Targ.Get_nRelations(iRel))
                                    then
                                    begin
                                        if iRel > 0 then
                                        begin
                                        RecursiveIA2Rel(ccNode, ia2Targ);
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

function TwndMSAAV.IA2Text4HTML(pAcc: IAccessible = nil;
	tab: string = ''): string;
var
	iSP: IServiceProvider;
	iInter: IInterface;
	ia2, ia2Targ: IAccessible2;
	iAV: IAccessibleValue;
	iaTarg: IAccessible;
	hr, i, iRole, t, p, iTarg, iUID: integer;
	MSAAs: array [0 .. 13] of WideString;
	Node: TTreeNode;
	ovChild, ovValue: OleVariant;
	iAL: IAccessibleRelation;
	s, ws: WideString;
	oList, cList: TStringList;
	iSOffset, IEOffset: integer;
	ia2Txt: IAccessibleText;

begin
	Result := '';
	if not Assigned(iAcc) then
		Exit;
	if pAcc = nil then
		pAcc := iAcc;
	// if (iMode <> 1) then Exit;
	if (not mnuMSAA.Checked) then
		Exit;
	oList := TStringList.Create;
	try

		hr := pAcc.QueryInterface(IID_IServiceProvider, iSP);
		if SUCCEEDED(hr) and Assigned(iSP) then
		begin
			try
				hr := iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE2, ia2);
			except
				hr := E_FAIL;
			end;
			if SUCCEEDED(hr) and Assigned(ia2) then
			begin
				for i := 0 to 13 do
				begin
					MSAAs[i] := None;
				end;

				ovChild := VarParent;
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
						if SUCCEEDED(ia2.Role(iRole)) then
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
									if cList[i] = '' then
										continue;

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
						if SUCCEEDED(ia2.Get_localizedExtendedStates(10, PWideString1(ws),
							iRole)) then
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
						if SUCCEEDED(iSP.QueryService(IID_IACCESSIBLE, iid_iaccessiblevalue,
							iAV)) then
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
				if (flgIA2 and TruncPow(2, 10)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.QueryInterface(IID_IAccessibleText, ia2Txt)) then
						begin

							ia2Txt.Get_attributes(1, iSOffset, IEOffset, ws);
							MSAAs[12] := SysUtils.StringReplace(ws, '\,', ',',
								[rfReplaceAll, rfIgnoreCase]);

						end;
					except
						on E: Exception do
							MSAAs[12] := E.Message;
					end;
				end;
				if (flgIA2 and TruncPow(2, 11)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.Get_uniqueID(iUID)) then
						begin

							ia2Txt.Get_attributes(1, iSOffset, IEOffset, ws);
							MSAAs[13] := inttostr(iUID);

						end;
					except
						on E: Exception do
							MSAAs[13] := E.Message;
					end;
				end;
				Result := Result + #13#10#9 + tab + '<strong>' + lIA2[0] + '</strong>' +
					#13#10#9 + '<ul>';
				for i := 0 to 11 do
				begin
					if (flgIA2 and TruncPow(2, i)) <> 0 then
					begin
						if i = 5 then
						begin
							Result := Result + #13#10#9#9 + tab + '<li>' + lIA2[i + 1] + ': '
								+ MSAAs[i] + '</li>';
						end
						else if (i = 4) then
						begin
							Node := nil;

							try
								if SUCCEEDED(ia2.Get_nRelations(iRole)) then
								begin
									for p := 0 to iRole - 1 do
									begin
										Result := Result + #13#10#9#9 + tab + '<li>' + rType + ': ';
										ia2.Get_relation(p, iAL);
										iAL.Get_RelationType(s);

										Result := Result + s + '</li>';
										Result := Result + #13#10#9#9 + tab + '<li>' + rTarg + ': ';

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
															if s = '' then
																s := None;
															// iaTarg.Get_accValue(0, s);
															ia2Targ.Role(iRole);
															ws := GetIA2Role(iRole);
															if ws = '' then
																ws := None;
															// iaTarg.Get_accValue(0, ws);
															Result := Result + s + ' - ' + ws;
															{ if t <> itarg then
																Result := Result + ','; }
														end;
													end;
												end;
											end;
										end;
										Result := Result + ';</li>';
									end;
								end;
							except
								on E: Exception do
									ShowErr(E.Message);
							end;
						end
						else if i = 9 then
						begin

							for t := 0 to 2 do
							begin
								Result := Result + #13#10#9#9 + tab + '<li>' + lIA2[11 + t] +
									': ' + MSAAs[t + 9] + ';</li>';
							end;

						end
						else if i = 10 then
						begin

							Result := Result + #13#10#9#9 + tab + '<li>' + lIA2[14] + ': ' +
								MSAAs[12] + '</li>';
						end
						else if i = 11 then
						begin

							Result := Result + #13#10#9#9 + tab + '<li>' + lIA2[15] + ': ' +
								MSAAs[13] + '</li>';
						end
						else
						begin

							Result := Result + #13#10#9#9 + tab + '<li>' + lIA2[i + 1] + ': '
								+ MSAAs[i] + ';</li>';
						end;
					end;
				end;


				// Result := Result + lIA2[9] + ':' + #9 + MSAAs[8] + #13#10;

			end;
		end;
	finally
		oList.Free;
	end;

end;

function TwndMSAAV.SetNodeData(pNode: PVirtualNode; Text1, Text2: string; Acc: iAccessible; iID: integer): PVirtualNode;
var
  ND: PNodeData;
begin
  result := TreeList1.InsertNode(pNode, amAddChildLast , nil);
  ND := TreeList1.GetNodeData(result);
  ND.Value1 := Text1;
  ND.Value2 := Text2;
  ND.Acc := Acc;
  ND.iID := iID;
end;

function TwndMSAAV.SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
var
	iSP: IServiceProvider;
	iInter: IInterface;
	ia2, ia2Targ: IAccessible2;
	iAV: IAccessibleValue;
	iaTarg: IAccessible;

	hr, i, iRole, t, p, iTarg, iPos, iUID: integer;
	MSAAs: array [0 .. 13] of WideString;
	rNode, Node, cNode, ccNode, cccNode: PVirtualNode;
	ovChild, ovValue: OleVariant;
	iAL: IAccessibleRelation;
	s, ws: WideString;
	oList, cList: TStringList;
	iSOffset, IEOffset: integer;
	ia2Txt: IAccessibleText;
begin
	Result := '';
	if not Assigned(iAcc) then
		Exit;
	if pAcc = nil then
		pAcc := iAcc;
	if (not mnuMSAA.Checked) then
		Exit;

	// if flgIA2 = 0 then Exit;
	rNode := nil;
	oList := TStringList.Create;
	try

		hr := pAcc.QueryInterface(IID_IServiceProvider, iSP);
		if SUCCEEDED(hr) and (Assigned(iSP)) then
		begin
			try
				hr := iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE2, ia2);

			except
				hr := E_FAIL;
			end;
			if SUCCEEDED(hr) and Assigned(ia2) then
			begin
				for i := 0 to 8 do
				begin
					MSAAs[i] := None;
				end;

				ovChild := VarParent;
				if (flgIA2 and TruncPow(2, 0)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.Get_accName(ovChild, ws)) then
						begin
							MSAAs[0] := ws;
						end;
					except
						on E: Exception do
						begin
							MSAAs[0] := E.Message;

						end;
					end;
				end;

				if (flgIA2 and TruncPow(2, 1)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.Role(iRole)) then
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
									if cList[i] = '' then
										continue;

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
						if SUCCEEDED(ia2.Get_localizedExtendedStates(10, PWideString1(ws),
							iRole)) then
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
						if SUCCEEDED(iSP.QueryService(iid_iaccessiblevalue,
							iid_iaccessiblevalue, iAV)) then
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

				if (flgIA2 and TruncPow(2, 10)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.QueryInterface(IID_IAccessibleText, ia2Txt)) then
						begin

							ia2Txt.Get_attributes(1, iSOffset, IEOffset, ws);
							MSAAs[12] := SysUtils.StringReplace(ws, '\,', ',',
								[rfReplaceAll, rfIgnoreCase]);

						end;
					except
						on E: Exception do
							MSAAs[12] := E.Message;
					end;
				end;

				if (flgIA2 and TruncPow(2, 11)) <> 0 then
				begin
					try
						if SUCCEEDED(ia2.Get_uniqueID(iUID)) then
						begin
							MSAAs[13] := inttostr(iUID);

						end;
					except
						on E: Exception do
							MSAAs[13] := E.Message;
					end;
				end;
				Result := lIA2[0] + #13#10;
				// nodes := TreeList1.Items;
				if SetTL then
				begin
					// rNode := nodes.AddChild(nil, lIA2[0]);
					rNode := SetNodeData(nil, lIA2[0], '', nil, 0);
				end;
				for i := 0 to 11 do
				begin
					if (flgIA2 and TruncPow(2, i)) <> 0 then
					begin
						if i = 5 then
						begin
							Result := Result + lIA2[i + 1] + ':' + #9 + MSAAs[i] + '(';
							// + #13#10;
							if SetTL then
							begin
								// node := nodes.AddChild(rNode, lIA2[i+1]);
								Node := SetNodeData(rNode, lIA2[i + 1], '', nil, 0);
								for p := 0 to oList.Count - 1 do
								begin
									t := pos(':', oList[p]);
									SetNodeData(Node, copy(oList[p], 0, t - 1),
										copy(oList[p], t + 1, Length(oList[p])), nil, 0);
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
						else if (i = 4) { and ((flgIA2 and 16) <> 0) } then
						begin
							try
								if SUCCEEDED(ia2.Get_nRelations(iRole)) then
								begin
									for p := 0 to iRole - 1 do
									begin
										Result := Result + #13#10 + #9 + rType + ':';
										ia2.Get_relation(p, iAL);
										iAL.Get_RelationType(s);

										ccNode := nil;

										if SetTL then
										begin
											Node := SetNodeData(rNode, lIA2[i + 1], '', nil, 0);
											cNode := SetNodeData(Node, inttostr(p), '', nil, 0);
											SetNodeData(cNode, rType, s, nil, 0);
											ccNode := SetNodeData(cNode, rTarg, '', nil, 0);
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
															if s = '' then
																s := None;
															// iaTarg.Get_accValue(0, s);
															ia2Targ.Role(iRole);
															ws := GetIA2Role(iRole);
															if ws = '' then
																ws := None;
															// iaTarg.Get_accValue(0, ws);
															Result := Result + s + ' - ' + ws + #13#10 + #9;
															if SetTL then
															begin
																cccNode := SetNodeData(ccNode, inttostr(t),
																	s + ' - ' + ws, iaTarg, 0);
															end;
														end;
													end;
												end;
											end;
										end;

									end;
								end;
							except
								on E: Exception do
									ShowErr(E.Message);
							end;
						end
						else if i = 9 then
						begin
							Node := nil;
							if SetTL then
							begin
								// node := nodes.AddChild(rNode, lIA2[10]);
								Node := SetNodeData(rNode, lIA2[10], '', nil, 0);
							end;
							for t := 0 to 2 do
							begin
								if SetTL then
								begin
									// cnode := nodes.AddChild(Node, lIA2[11+t]);
									// TreeList1.SetNodeColumn(cnode, 1, MSAAs[t+9]);
									SetNodeData(Node, lIA2[11 + t], MSAAs[t + 9], nil, 0);
								end;
								Result := Result + lIA2[11 + t] + ':' + #9 +
									MSAAs[t + 9] + #13#10;
							end;

						end
						else if i = 10 then
						begin

							if SetTL then
							begin
								// Node := nodes.AddChild(rNode, lIA2[14]);

								if MSAAs[12] = '' then
								begin
									SetNodeData(rNode, lIA2[14], None, nil, 0);
									// TreeList1.SetNodeColumn(Node, 1, None);
								end
								else
								begin
									Node := SetNodeData(rNode, lIA2[14], '', nil, 0);
									cList := TStringList.Create;
									try
										cList.Delimiter := ';';

										cList.StrictDelimiter := True;
										cList.DelimitedText := MSAAs[12];
										for p := 0 to cList.Count - 1 do
										begin
											if cList[p] <> '' then
											begin
												iPos := pos(':', cList[p]);
												SetNodeData(Node, copy(cList[p], 0, iPos - 1),
													copy(cList[p], iPos + 1, Length(cList[p])), nil, 0);
												// cNode := nodes.AddChild(Node, copy(cList[p], 0, iPos - 1));
												// TreeList1.SetNodeColumn(cNode, 1, copy(cList[p], iPos + 1, length(cList[p])));
											end;
										end;

									finally
										cList.Free;
									end;
								end;
							end;
							Result := Result + lIA2[14] + ':' + #9 + MSAAs[12] + #13#10;
						end
						else if i = 11 then
						begin
							if SetTL then
							begin
								// node := nodes.AddChild(rNode, lIA2[i+1]);
								// TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
								SetNodeData(rNode, lIA2[15], MSAAs[13], nil, 0);
							end;
							Result := Result + lIA2[15] + ':' + #9 + MSAAs[13] + #13#10;
						end
						else
						begin
							if SetTL then
							begin
								// node := nodes.AddChild(rNode, lIA2[i+1]);
								// TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
								SetNodeData(rNode, lIA2[i + 1], MSAAs[i], nil, 0);
							end;
							Result := Result + lIA2[i + 1] + ':' + #9 + MSAAs[i] + #13#10;
						end;
					end;
				end;

				// Result := Result + lIA2[9] + ':' + #9 + MSAAs[8] + #13#10;
				if SetTL then
					TreeList1.Expanded[TreeList1.RootNode] := True;
				// rNode.Expand(True);
			end
			else // IAccessible2 failed
			begin
				if SetTL then
				begin
					rNode := SetNodeData(nil, lIA2[0],
						Err_Inter + '(IAccessible2)', nil, 0);
				end;
			end;
		end;
		TipTextIA2 := Result + #13#10 + #13#10;
		TreeList1.Expanded[rNode] := True;
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

        Result := Result + #13#10 + tab + #9 + '<strong>' + lMSAA[0] + '</strong>' + #13#10#9 + tab + '<ul>';
        for i := 0 to 10 do
        begin
            if (flgMSAA and TruncPow(2, i)) <> 0 then
            begin
                Result := Result + #13#10#9#9 + tab + '<li>' +  lMSAA[i+1] + ': ' + MSAAs[i] + ';</li>';
            end;
        end;
        Result := result + #13#10#9 + tab + '</ul>';
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
    ini := TMemIniFile.Create(TransPath, TEncoding.UTF8);
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

function TwndMSAAV.MSAAText4UIA(TextOnly: boolean = false): string;
var
	PC: PChar;
	ovValue, ovProp: OleVariant;
	MSAAs: array [0 .. 10] of WideString;
	ws: WideString;
	hr: HResult;
	Node: PVirtualNode;
	i: integer;
	cVal: cardinal;
	pEle: IUIAUTOMATIONELEMENT;
	arEle: IUIAutomationElementArray;
	pCond: IUIAutomationCondition;
	iVW: IUIAutomationTreeWalker;
	Scope: TreeScope;
	iLeg: IUIAutomationLegacyIAccessiblePattern;
	iInt: IInterface;
	function GetiLegacy: HResult;
	begin
  	result := S_FALSE;
  	hr := uiEle.GetCurrentPattern(UIA_LegacyIAccessiblePatternId, iInt);
		if (hr = 0) and (Assigned(iInt)) then
		begin
    	result := iInt.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern, iLeg);
		end;

	end;
begin
  if not mnuMSAA.Checked then
		Exit;
	if not assigned(uiEle) then
	begin
		Exit;
	end;

  hr := GetiLegacy;
	if (hr = 0) and (Assigned(iLeg)) then
	begin

		for i := 0 to 10 do
		begin
			MSAAs[i] := None;
		end;
		if (flgMSAA and 1) <> 0 then
		begin
			try
				iLeg.Get_CurrentName(ws);
        MSAAs[0] := ws;
			except
				on E: Exception do
					MSAAs[0] := E.Message;
			end;
		end;
		if (flgMSAA and 2) <> 0 then
		begin
			try
				iLeg.Get_CurrentRole(cVal);
        PC := StrAlloc(255);
        GetRoleTextW(cVal, PC, StrBufSize(PC));
        MSAAs[1] := PC;
        StrDispose(PC);
			except
				on E: Exception do
					MSAAs[1] := E.Message;
			end;
		end;
		if (flgMSAA and 4) <> 0 then
		begin
			try
        iLeg.Get_CurrentState(cVal);
        MSAAs[2] := GetMultiState(cVal);
			except
				on E: Exception do
					MSAAs[2] := E.Message;
			end;
		end;
		if (flgMSAA and 8) <> 0 then
		begin
			try
      	iLeg.Get_CurrentDescription(ws);
				MSAAs[3] := ws;
			except
				on E: Exception do
					MSAAs[3] := E.Message;
			end;
		end;

		if (flgMSAA and 16) <> 0 then
		begin
			try
      	iLeg.Get_CurrentDefaultAction(ws);
				MSAAs[4] := ws;
			except
				on E: Exception do
					MSAAs[4] := E.Message;
			end;
		end;

		if (flgMSAA and 32) <> 0 then
		begin
			try
      	iLeg.Get_CurrentValue(ws);
				MSAAs[5] := ws;
			except
				on E: Exception do
					MSAAs[5] := E.Message;
			end;
		end;

		if (flgMSAA and 64) <> 0 then
		begin
			try
				if SUCCEEDED(UIAuto.Get_ControlViewWalker(iVW)) then
				begin
					if (Assigned(iVW)) and (SUCCEEDED(iVW.GetParentElement(uiEle, pEle)))
					then
					begin
						if Assigned(pEle) and
							SUCCEEDED(pEle.GetCurrentPropertyValue
							(UIA_LegacyIAccessibleNamePropertyId, ovValue)) then
							MSAAs[6] := VarToStr(ovValue);
					end;
					iVW := nil;
				end;

			except
				on E: Exception do
					MSAAs[6] := E.Message;
			end;
		end;

		if (flgMSAA and 128) <> 0 then
		begin
			try
				i := 0;
				TVariantArg(ovProp).vt := VT_BOOL;
				TVariantArg(ovProp).vbool := True;
				UIAuto.CreatePropertyCondition(UIA_IsControlElementPropertyId,
					ovProp, pCond);
				Scope := TreeScope_Children;
				if SUCCEEDED(uiEle.FindAll(Scope, pCond, arEle)) then
				begin
					arEle.Get_Length(i);
				end;
				MSAAs[7] := inttostr(i);

			except
				on E: Exception do
					MSAAs[7] := E.Message;
			end;
		end;

		if (flgMSAA and 256) <> 0 then
		begin
			try
      	iLeg.Get_CurrentHelp(ws);
				MSAAs[8] := ws;
			except

				on E: Exception do
					MSAAs[8] := E.Message;
			end;
		end;
		if (flgMSAA and 512) <> 0 then
		begin

			MSAAs[9] := '';
		end;
		if (flgMSAA and 1024) <> 0 then
		begin
			try
      	iLeg.Get_CurrentKeyboardShortcut(ws);
				MSAAs[10] := ws;
			except

				on E: Exception do
					MSAAs[10] := E.Message;
			end;
		end;

		Result := lMSAA[0] + '(UIA_LegacyIAccessible)' + #13#10;
		if not TextOnly then
		begin
			// nodes := TreeList1.Items;
			// refNode := nodes.AddChild(nil, lMSAA[0]);
			refVNode := SetNodeData(nil, lMSAA[0] + '(UIA_LegacyIAccessible)',
				'', nil, 0);
			for i := 0 to 10 do
			begin
				if (flgMSAA and TruncPow(2, i)) <> 0 then
				begin
					// node := nodes.AddChild(refNode, lMSAA[i+1]);
					// TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
					Node := SetNodeData(refVNode, lMSAA[i + 1], MSAAs[i], nil, 0);
					Result := Result + lMSAA[i + 1] + ':' + #9 + MSAAs[i] + #13#10;
					if i = 4 then
					begin
						iDefIndex := Node.Index;
					end;
				end;
			end;
			// refNode.Expand(True);
			TreeList1.Expanded[refVNode] := True;
		end
		else
		begin
			for i := 0 to 10 do
			begin
				if (flgMSAA and TruncPow(2, i)) <> 0 then
				begin
					Result := Result + lMSAA[i + 1] + ':' + #9 + MSAAs[i] + #13#10;
				end;
			end;
		end;
		TipText := Result + #13#10;
	end
  else
  begin
  	SetNodeData(nil, lMSAA[0], Err_Inter + '(IUIAutomationLegacyIAccessiblePattern)', nil, 0);
  end;
end;



function TwndMSAAV.MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = false): string;
var
    PC:PChar;
    ovValue, ovChild: OleVariant;
    MSAAs: array [0..10] of widestring;
    ws: widestring;
    node: pVirtualNode;
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
        	//nodes := TreeList1.Items;
        	//refNode := nodes.AddChild(nil, lMSAA[0]);
          refVNode := SetNodeData(nil, lMSAA[0], '', nil, 0);
        	for i := 0 to 10 do
        	begin
            	if (flgMSAA and TruncPow(2, i)) <> 0 then
            	begin
                	//node := nodes.AddChild(refNode, lMSAA[i+1]);
                	//TreeList1.SetNodeColumn(node, 1, MSAAs[i]);
                  Node := SetNodeData(refVNode, lMSAA[i+1], MSAAs[i], nil, 0);
                	Result := Result + lMSAA[i+1] + ':' + #9 + MSAAs[i] + #13#10;
                	if i = 4 then
                	begin
                    	iDefIndex := Node.Index;
                	end;
            	end;
        	end;
          TreeList1.Expanded[refVNode] := True;
          //refNode.Expand(True);
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
	i: integer;
	rNode: PVirtualNode;
begin
	if not Assigned(CEle) then
	begin
		//GetNaviState(True);

		Exit;
	end;
	for i := 0 to 2 do
		HTMLs[i, 1] := '';
		rNode := SetNodeData(nil, sHTML, '', nil, 0);


    if Assigned(HTMLTH) then
		begin
			HTMLTH.Terminate;
			HTMLTH.WaitFor;
			HTMLTH.Free;
			HTMLTH := nil;
		end;

    HTMLTH := HTMLThread.Create(CEle, rNode);
    HTMLTH.OnTerminate := THHTMLDone;
    HTMLTH.Start;


end;

function TwndMSAAV.HTMLText4FF: string;
var

    hAttrs, s, Path, d: string;
    i: integer;
    aPC, aPC2: array [0..64] of pWidechar;
    aSI: array [0..64] of Smallint;
    PC, PC2:PChar;
    SI: Smallint;
    PU, PU2: PUINT;
    WD, tWD: Word;
    rNode, Node, sNode: pVirtualNode;
    iText: ISimpleDOMText;
    iSP: iServiceProvider;
begin
    if not Assigned(SDOM) then
    begin
        //GetNaviState(True);
        Exit;
    end;
    for i := 0 to 2 do HTMLsFF[i, 1] := '';
    try
        Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
        SDOM.get_nodeInfo(PC, SI, PC2, PU, PU2, tWD);

        if tWD <> 3 then
            HTMLsFF[0, 1] := PC
        else
            HTMLsFF[0, 1] := sTxt;
        //rNode := TreeList1.Items.AddChild(nil, sHTML);
        rNode := SetNodeData(nil, sHTML, '', nil, 0);
        //Node := TreeList1.Items.AddChild(rNode, HTMLsFF[0, 0]);
        //TreeList1.SetNodeColumn(node, 1, HTMLsFF[0, 1]);
        SetNodeData(rNode, HTMLsFF[0, 0], HTMLsFF[0, 1], nil, 0);
        SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
        if WD > 0 then
        begin
            //Node := TreeList1.Items.AddChild(rNode, HTMLsFF[1, 0]);
            Node := SetNodeData(rNode, HTMLsFF[1, 0], '', nil, 0);
            sNode := Node;

            for i := 0 to WD - 1 do
            begin
                s := LowerCase(aPC[i]);
                if (s <> 'role') and (Copy(s, 1, 4) <> 'aria') then
                begin
                    try

                        //Node := TreeList1.Items.AddChild(sNode, aPC[i]);
                        //TreeList1.SetNodeColumn(node, 1, aPC2[i]);
                        SetNodeData(sNode, aPC[i], aPC2[i], nil, 0);
                        hattrs := hattrs + WideString(aPC[i]) + '","' + WideString(aPC2[i]) + #13#10;
                    except

                    end;
                end;
            end;
        end
        else
        begin
            //Node := TreeList1.Items.AddChild(rNode, HTMLsFF[1, 0]);
            //TreeList1.SetNodeColumn(node, 1, none)
            SetNodeData(rNode, HTMLsFF[1, 0], none, nil, 0);
        end;
        if hattrs = '' then
            HTMLsFF[1, 1] := none
        else
            HTMLsFF[1, 1] := hattrs;
        PC := '';
        if tWD <> 3 then
        begin
            SDOM.get_innerHTML(PC);
            d := HTMLsFF[2, 0] + StypeFF;
        end
        else
        begin
            iSP := SDOM as IServiceProvider;
            if SUCCEEDED(iSP.QueryInterface(IID_ISIMPLEDOMTEXT, iText)) then
            begin
                iText.get_domText(PC);
                d := HTMLsFF[2, 0] + '(' + sTxt + ')';
            end;
        end;

        HTMLsFF[2, 1] := PC;

        Result := sHTML + #13#10 + HTMLsFF[0, 0] + ':' + #9 +  HTMLsFF[0, 1] +
                  #13#10 + HTMLsFF[1, 0] + ':' + #9 +  HTMLsFF[1, 1] +
                  #13#10 + d + ':' + #9 +  HTMLsFF[2, 1] + #13#10#13#10;
        Memo1.Text := d + ':' + #13#10 +  HTMLsFF[2, 1] + #13#10#13#10;
        //rNode.Expand(True);
        TreeList1.Expanded[rNode] := true;
    except

    end;
end;

procedure Recursive(sID: string; iChild: integer; ISEle:ISimpleDOMNODE; var outEle: ISimpleDOMNODE);
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
            Recursive(sID, Integer(pChild), ChildNode, outEle);
        end;
    end;
end;

function TwndMSAAV.GetSameAcc: boolean;
var
    i: integer;
    rNode: TTreeNode;
    cAccName: widestring;
    cAccRole: string;
    ov: olevariant;
begin
    Result := false;
    try
      ov := varParent;
      iAcc.Get_accName(ov, cAccName);
    	if cAccName = '' then
        cAccName := None;
    	cAccName := StringReplace(cAccName, #13, ' ', [rfReplaceAll]);
    	cAccName := StringReplace(cAccName, #10, ' ', [rfReplaceAll]);
    	cAccRole := Get_RoleText(iAcc, integer(varParent));
      cAccName := cAccName + ' - ' + cAccRole;
      if TBList.Count > 0 then
      begin
      	for i := 0 to TBList.Count - 1 do
        begin
        	Application.ProcessMessages;
        	rNode := TreeView1.Items.GetNode(HTreeItem(TBList.Items[i]));
        	if Assigned(rNode) and (rNode.Text = cAccName) then
    			begin
    				if IsSameUIElement(iAcc, TTreeData(rNode.Data^).Acc, varParent, TTreeData(rNode.Data^).iID) then
						begin
							TreeView1.OnChange := nil;
							rNode.Expanded := True;
							rNode.Selected := True;
							Result := True;
							if (acShowTip.Checked) then
							begin
								ShowTipWnd;
							end;
              GetNaviState;
							break;
						end;
    			end;
        end;

      end;

    finally
      TreeView1.OnChange := TreeView1Change;
    end;
  end;

procedure TwndMSAAV.RecursiveID(sID: string; isEle: ISimpleDOMNODE; Labelled: boolean = true);
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
        Recursive(sID, Integer(iChild), ISEle, IDEle);
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

	iAttrCol: IHTMLATTRIBUTECOLLECTION;
	iDomAttr: IHTMLDOMATTRIBUTE;
	aEle: IHTMLElement;
	s: string;
	List, List2: TStringList;
	ovValue: OleVariant;
	i: integer;
	aAttrs, LowerS: string;
	rNode, Node, rcNode: PVirtualNode;
	aPC, aPC2: array [0 .. 64] of PChar;
	aSI: array [0 .. 64] of Smallint;
	WD: Word;
	Serv: IServiceProvider;
	lAcc: IAccessible;
	RC: TRect;
begin

    if (CEle = nil) or (SDom = nil) then
	begin
		// GetNaviState(True);
		Exit;
	end;
	if Assigned(WndLabel) then
	begin
		FreeAndNil(WndLabel);
		WndLabel := nil;
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

	if (Assigned(CEle)) then
	begin
		iAttrCol := (CEle as IHTMLDomNode).attributes as IHTMLATTRIBUTECOLLECTION;
		for i := 0 to iAttrCol.Length - 1 do
		begin
			ovValue := i;
			iDomAttr := iAttrCol.item(ovValue) as IHTMLDOMATTRIBUTE;
			if iDomAttr.specified then
			begin
				s := LowerCase(VarToStr(iDomAttr.nodeName));
				if s = 'role' then
				begin
					try
						if VarHaveValue(iDomAttr.nodeValue) then
							ARIAs[0, 1] := VarToStr(iDomAttr.nodeValue)
						else
							ARIAs[0, 1] := ''
					except
						on E: Exception do
						begin
							ShowErr(E.Message);
							ARIAs[0, 1] := ''
						end;
					end;
				end
				else if copy(s, 1, 4) = 'aria' then
				begin

					try
						if VarHaveValue(iDomAttr.nodeValue) then
						begin

							aAttrs := aAttrs + '"' + VarToStr(iDomAttr.nodeName) + '","' +
								(iDomAttr.nodeValue) + '"' + #13#10;
							LowerS := LowerCase(VarToStr(iDomAttr.nodeName));
							if acRect.Checked then
							begin
								if LowerS = 'aria-labelledby' then
								begin
									aEle := (CEle.Document as IHTMLDOCUMENT3)
										.getElementById(VarToStr(iDomAttr.nodeValue));
									if Assigned(aEle) then
									begin
										if SUCCEEDED(aEle.QueryInterface(IID_IServiceProvider, Serv))
										then
										begin
											if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE,
												IID_IACCESSIBLE, lAcc)) then
											begin
												// if SUCCEEDED(lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0)) then
												// begin
												lAcc.accLocation(RC.left, RC.top, RC.right,
													RC.bottom, 0);
												ShowLabeledWnd(clBlue, RC);
												// end;

											end;
										end;
									end;
								end
								else if LowerS = 'aria-decirbedby' then
								begin
									aEle := (CEle.Document as IHTMLDOCUMENT3)
										.getElementById(VarToStr(iDomAttr.nodeValue));
									if Assigned(aEle) then
									begin
										if SUCCEEDED(aEle.QueryInterface(IID_IServiceProvider, Serv))
										then
										begin
											if SUCCEEDED(Serv.QueryService(IID_IACCESSIBLE,
												IID_IACCESSIBLE, lAcc)) then
											begin
												// if SUCCEEDED(lAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, 0)) then
												// begin
												lAcc.accLocation(RC.left, RC.top, RC.right,
													RC.bottom, 0);
												ShowDescWnd(clBlue, RC);
												// end;
											end;

										end;
									end;
								end;
							end;
						end;

					except
						on E: Exception do
						begin
							ShowErr(E.Message);
							aAttrs := aAttrs + '';
						end;
					end;
				end;
			end;
		end;
	end
	else if (Assigned(SDom)) then
	begin
		// function get_attributes(maxAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR; out numAttribs: WORD): HRESULT; stdcall;
		// function get_attributesForNames(numAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR): HRESULT; stdcall;
		SDom.Get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
		for i := 0 to WD - 1 do
		begin
			LowerS := LowerCase(aPC[i]);
			if LowerS = 'role' then
			begin
				ARIAs[0, 1] := aPC2[i];
			end
			else if copy(LowerS, 1, 4) = 'aria' then
			begin

				aAttrs := aAttrs + '"' + aPC[i] + '","' + aPC2[i] + '"' + #13#10;
				if acRect.Checked then
				begin
					if LowerS = 'aria-labelledby' then
					begin
						// showmessage(aPC2[i]);
						RecursiveID(aPC2[i], SDom, True);
					end
					else if LowerS = 'aria-decirbedby' then
						RecursiveID(aPC2[i], SDom, false);
				end;
			end;
		end;
	end;
	if ARIAs[0, 1] = '' then
		ARIAs[0, 1] := None;
	if aAttrs = '' then
		ARIAs[1, 1] := None
	else
		ARIAs[1, 1] := aAttrs;

	{ rNode := TreeList1.Items.AddChild(nil, sAria);
		Node := TreeList1.Items.AddChild(rNode, Arias[0, 0]);
		TreeList1.SetNodeColumn(node, 1, Arias[0, 1]); }
	rNode := SetNodeData(nil, sARIA, '', nil, 0);
	Node := SetNodeData(rNode, ARIAs[0, 0], ARIAs[0, 1], nil, 0);

	// Node := TreeList1.Items.AddChild(rNode, Arias[1, 0]);
	if aAttrs = '' then
	begin
		// TreeList1.SetNodeColumn(node, 1, Arias[1, 1]);
		SetNodeData(rNode, ARIAs[1, 0], ARIAs[1, 1], nil, 0);
	end
	else
	begin
		rcNode := Node;
		List := TStringList.Create;
		List2 := TStringList.Create;
		try
			List.Text := aAttrs;
			for i := 0 to List.Count - 1 do
			begin
				List2.Clear;
				List2.CommaText := List[i];
				if List2.Count >= 1 then
				begin
					// Node := TreeList1.Items.AddChild(rcNode, List2[0]);
					if List2.Count >= 2 then
					begin
						// TreeList1.SetNodeColumn(node, 1, List2[1]);
						SetNodeData(rcNode, List2[0], List2[1], nil, 0);
					end
					else
						SetNodeData(rcNode, List2[0], '', nil, 0);
				end;
			end;
		finally
			List.Free;
			List2.Free;
		end;
	end;
	Result := sARIA + #13#10 + ARIAs[0, 0] + ':' + #9 + ARIAs[0, 1] + #13#10 +
		ARIAs[1, 0] + ':' + #9 + ARIAs[1, 1] + #13#10#13#10;
	// rNode.Expand(True);
	TreeList1.Expanded[rNode] := True;
end;

function TwndMSAAV.IsSameUIElement(ia1, ia2: IAccessible; iID1, iID2: integer): boolean;
var
    UIEle1, UIEle2: IUIAutomationElement;
    iSame: integer;
    hr: hresult;
    tagPT: UIAutomationClient_TLB.tagPoint;
    cRC: TRect;
begin
	Result := false;
	if Assigned(UIAuto) and Assigned(ia1) and Assigned(ia2) then
	begin
  	hr := ia1.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, iID1);
    tagPT.X := cRC.Location.X;
		tagPT.Y := cRC.Location.Y;
		if (hr = 0) and SUCCEEDED(UIAuto.ElementFromPoint(tagPT, UIEle1)) then
		begin
      hr := ia2.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, iID2);
    	tagPT.X := cRC.Location.X;
			tagPT.Y := cRC.Location.Y;
			if (hr = 0) and SUCCEEDED(UIAuto.ElementFromPoint(tagPT, UIEle2)) then
			begin
				hr := UIAuto.CompareElements(UIEle1, UIEle2, iSame);
				if SUCCEEDED(hr) and (iSame <> 0) then
					Result := True;
			end;
		end;

	end;
end;


function TwndMSAAV.UIAText(HTMLout: boolean = false; tab: string = ''): string;
var
	iRes, i, t, iLen, iBool, iStyleID: integer;
	dblRV: double;
	rNode, Node, cNode: PVirtualNode;
	ini: TMemInifile;
	ResEle: IUIAUTOMATIONELEMENT;
	UIArray: IUIAutomationElementArray;
	iaTarg: IAccessible;
	oVal: OleVariant;
	ws: WideString;
	RC: tagRECT;
	p: Pointer;
	OT: OrientationType;
	iInt: IInterface;
	UIAs: array [0 .. 59] of string;
	iLeg: IUIAutomationLegacyIAccessiblePattern;
	iRV: IUIAutomationRangeValuePattern;
	iSID: IUIAutomationStylesPattern;
	iSP: IStylesProvider;
	iIface: IInterface;
	hr: HResult;
	ND: PNodeData;

  tagPT: UIAutomationClient_TLB.tagPoint;
  cRC: TRect;
const
	IsProps: Array [0 .. 19] of integer = (30027, 30028, 30029, 30030, 30031,
		30032, 30033, 30034, 30035, 30036, 30037, 30038, 30039, 30040, 30041, 30042,
		30043, 30044, 30108, 30109);

	function GetMSAA: IAccessible;
	var
		rIacc: IAccessible;
	begin
		Result := nil;
		if SUCCEEDED(ResEle.GetCurrentPattern(10018, iInt)) then
		begin
			if SUCCEEDED(iInt.QueryInterface
				(IID_IUIAutomationLegacyIAccessiblePattern, iLeg)) then
			begin
				if SUCCEEDED(iLeg.GetIAccessible(rIacc)) then
					Result := rIacc;
			end;
		end;

	end;
	function PropertyPatternIS(iID: integer): string;
	begin

		case iID of
			30001:
				Result := 'rect';
			30002, 30012:
				Result := 'int';
			30003:
				Result := 'controltypeid';
			30008, 30009, 30010, 30016, 30017, 30019, 30022, 30025, 30103:
				Result := 'bool';
			// 30018: Result := 'iuiautomationelement';
			30020:
				Result := 'int'; // 'uia_hwnd';
			// 30104, 30105, 30106: Result := 'iuiautomationelementarray';
			30023:
				Result := 'orientationtype';
		else
			Result := 'str';
		end;
	end;
	function GetCID(iID: integer): string;
	begin
		ini := TMemInifile.Create(TransPath, TEncoding.UTF8);
		try
			Result := ini.ReadString('UIA', inttostr(iID), None);
		finally
			ini.Free;
		end;
	end;
	function GetOT: string;
	begin
		ini := TMemInifile.Create(TransPath, TEncoding.UTF8);
		try
			if OT = OrientationType_Horizontal then
				Result := ini.ReadString('UIA', 'OrientationType_Horizontal',
					'Horizontal')
			else if OT = OrientationType_Vertical then
				Result := ini.ReadString('UIA', 'OrientationType_Vertical', 'Vertical')
			else
				Result := ini.ReadString('UIA', 'OrientationType_None', 'None');
		finally
			ini.Free;
		end;
	end;

begin
	if not Assigned(UIAuto) then
		Exit;
	if not mnutvUIA.Checked then
	begin
		uiEle := nil;
    iAcc.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, varParent);
    tagPT.X := cRC.Location.X;
		tagPT.Y := cRC.Location.Y;
		UIAuto.ElementFromPoint(tagPT, uiEle);
	end;
	if (not Assigned(uiEle)) then
		Exit;

	try
		// iRes := UIA.ElementFromIAccessible(iAcc, 0, UIEle);
		if { (iRes = S_OK) and } (Assigned(uiEle)) then
		begin
			for i := 0 to 32 do
				UIAs[i] := None;

			if (flgUIA and TruncPow(2, 0)) <> 0 then
			begin
				try
					if SUCCEEDED(uiEle.Get_CurrentAcceleratorKey(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentAccessKey(ws)) then
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

					if SUCCEEDED(uiEle.Get_CurrentAriaProperties(ws)) then
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
					uiEle.Get_CurrentAriaRole(ws);
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
					if SUCCEEDED(uiEle.Get_CurrentAutomationId(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentBoundingRectangle(RC)) then
					begin
						UIAs[5] := inttostr(RC.left) + ',' + inttostr(RC.top) + ',' +
							inttostr(RC.right) + ',' + inttostr(RC.bottom);
					end;
				except
					on E: Exception do
						UIAs[5] := E.Message;
				end;
			end;
			if (flgUIA and TruncPow(2, 6)) <> 0 then
			begin
				try
					if SUCCEEDED(uiEle.Get_CurrentClassName(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentControlType(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentCulture(i)) then
					begin
						UIAs[8] := inttostr(i);
					end;
				except
					on E: Exception do
						UIAs[8] := E.Message;
				end;
			end;
			if (flgUIA and TruncPow(2, 9)) <> 0 then
			begin
				try
					if SUCCEEDED(uiEle.Get_CurrentFrameWorkID(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentHasKeyboardFocus(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentHelpText(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsControlElement(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsContentElement(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsDataValidForForm(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsEnabled(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsKeyboardFocusable(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsOffscreen(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsPassword(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentIsRequiredForForm(i)) then
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
					if SUCCEEDED(uiEle.Get_CurrentItemStatus(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentItemType(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentLocalizedControlType(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentName(ws)) then
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
					if SUCCEEDED(uiEle.Get_CurrentNativeWindowHandle(p)) then
					begin
						UIAs[24] := inttostr(integer(p));
					end;
				except
					on E: Exception do
						UIAs[24] := E.Message;
				end;
			end;
			if (flgUIA and TruncPow(2, 25)) <> 0 then
			begin
				try
					if SUCCEEDED(uiEle.Get_CurrentOrientation(OT)) then
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
					if SUCCEEDED(uiEle.Get_CurrentProcessId(i)) then
					begin
						UIAs[26] := inttostr(i);
					end;
				except
					on E: Exception do
						UIAs[26] := E.Message;
				end;
			end;
			if (flgUIA and TruncPow(2, 27)) <> 0 then
			begin
				try
					if SUCCEEDED(uiEle.Get_CurrentProviderDescription(ws)) then
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
				try // LiveSetting
					if SUCCEEDED(uiEle.GetCurrentPropertyValue(30135, oVal)) then
					begin
						if VarHaveValue(oVal) then
						begin
							if VarIsStr(oVal) or VarIsNumeric(oVal) then
								UIAs[32] := oVal;
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
					try // LiveSetting
						if SUCCEEDED(uiEle.GetCurrentPropertyValue(IsProps[i], oVal)) then
						begin
							if VarHaveValue(oVal) then
							begin
								if VarIsType(oVal, varBoolean) then
								begin
									iBool := oVal;
									UIAs[33 + i] := IfThen(iBool <> 0, sTrue, sFalse);
								end;
							end;
						end;
					except
						on E: Exception do
							UIAs[33 + i] := E.Message;
					end;
				end;
			end;
			if (flgUIA2 and TruncPow(2, 22)) <> 0 then
			begin
				// #30047~30052
				// GetCurrentPatternAs http://msdn.microsoft.com/en-us/library/windows/desktop/ee696039(v=vs.85).aspx
				// UIA_RangeValuePatternId 10003
				try
					hr := uiEle.GetCurrentPattern(UIA_RangeValuePatternId, iIface);
					if hr = S_OK then
					begin
						if Assigned(iIface) then
						begin
							if SUCCEEDED
								(iIface.QueryInterface(IID_IUIAutomationRangeValuePattern, iRV))
							then
							begin
								if SUCCEEDED(iRV.Get_currentValue(dblRV)) then
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
			end;
			if (flgUIA2 and TruncPow(2, 23)) <> 0 then
			begin
				try
					hr := uiEle.GetCurrentPattern(UIA_StylesPatternId, iIface);
					if (hr = S_OK) and Assigned(iIface) then
					begin
						hr := iIface.QueryInterface(IID_IUIAutomationStylesPattern, iSID);
						if (hr = S_OK) and Assigned(iSP) then
						// hr := iiFace.QueryInterface(IID_IStylesProvider, iSP);
						// if (hr = S_OK) and Assigned(iSP) then
						begin
							hr := iSID.Get_CurrentStyleId(iStyleID);
							// outputdebugstring(PWideChar('current style id is ' + inttostr(istyleid)));
							// hr := iSP.get_StyleId(iStyleID);
							if (hr = S_OK)
							{ and ((iStyleID >= 70000) and (iStyleID <= 70016)) } then
							begin
								// UIAs[59] := StyleID[iStyleID - 70000];
								UIAs[59] := inttostr(iStyleID);
							end;
						end;
					end;
				except
					on E: Exception do
						UIAs[59] := E.Message;
				end;
			end;
      if HTMLout then
      	Result := Result + #13#10 + tab + #9 + '<strong>' + lUIA[0] + '</strong>' + #13#10#9 + tab + '<ul>'
      else
				Result := lUIA[0] + #13#10;

      if not HTMLout then
				rNode := SetNodeData(nil, lUIA[0], '', nil, 0);
			for i := 0 to 54 do
			begin

				if i < 31 then
				begin
					if (flgUIA and TruncPow(2, i)) <> 0 then
					begin
						if i >= 28 then
						begin
							// node := nodes.AddChild(rNode, lUIA[i+1]);
              if not HTMLout then
								Node := SetNodeData(rNode, lUIA[i + 1], '', nil, 0);
							UIArray := nil;
							iRes := E_FAIL;
							if i = 28 then
							begin
								iRes := uiEle.Get_CurrentControllerFor(UIArray);
							end
							else if i = 29 then
							begin
								iRes := uiEle.Get_CurrentDescribedBy(UIArray);
							end
							else if i = 30 then
							begin
								iRes := uiEle.Get_CurrentFlowsTo(UIArray);
							end;
							if Assigned(UIArray) and SUCCEEDED(iRes) then
							begin
								if SUCCEEDED(UIArray.Get_Length(iLen)) then
								begin
									for t := 0 to iLen - 1 do
									begin
										// cNode := nodes.AddChild(Node, inttostr(t));
                    if not HTMLout then
											cNode := SetNodeData(Node, inttostr(t), '', nil, 0);
										ResEle := nil;
										if SUCCEEDED(UIArray.GetElement(t, ResEle)) then
										begin
											if Assigned(ResEle) then
											begin
												ws := None;
												if (not mnutvUIA.Checked) and (not HTMLOut) then
												begin
													iaTarg := GetMSAA;
													if Assigned(iaTarg) then
													begin
														iaTarg.Get_accName(0, ws);
														ND := TreeList1.GetNodeData(cNode);
														ND.Acc := iaTarg;
														ND.iID := 0;

													end;
												end
												else
												begin
													// UIA_LegacyIAccessibleNamePropertyId = 30092
													ResEle.GetCurrentPropertyValue(30092, oVal);
													ws := string(oVal);
												end;
                        if not HTMLout then
                        begin
													ND := TreeList1.GetNodeData(cNode);
													ND.Value2 := ws;
													// TreeList1.SetNodeColumn(cnode, 1, ws);
													Result := Result + lUIA[i + 1] + ':' + #9 + ws + #13#10;
                        end
                        else
                        begin
                        	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[i+1] + ': ' + ws + ';</li>';
                        end;
											end;
										end;
									end;
									TreeList1.Expanded[Node] := True;
								end;
							end
							else
							begin
								if not HTMLout then
                begin
									ND := TreeList1.GetNodeData(Node);
									ND.Value2 := None;
                end;
							end;
						end
						else
						begin
							if not HTMLout then
              begin
								SetNodeData(rNode, lUIA[i + 1], UIAs[i], nil, 0);
								Result := Result + lUIA[i + 1] + ':' + #9 + UIAs[i] + #13#10;
              end
              else
              begin
              	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[i + 1]  + ': ' + UIAs[i] + ';</li>';
              end;
						end;
					end;
				end
				else
				begin
					if (flgUIA2 and TruncPow(2, i - 31)) <> 0 then
					begin
						if i = 31 then
						begin
							if not HTMLout then
              begin
								Node := SetNodeData(rNode, lUIA[i + 1], '', nil, 0);
								ND := TreeList1.GetNodeData(Node);
              end;
							try
								iRes := uiEle.Get_CurrentLabeledBy(ResEle);
								if SUCCEEDED(iRes) then
								begin

									if not SUCCEEDED(ResEle.Get_CurrentName(ws)) then
										ws := None;

                  if not HTMLout then
									begin
										ND.Value2 := ws;
										if not mnutvUIA.Checked then
										begin
											ND.Acc := GetMSAA;
											ND.iID := 0;
										end;
										Result := Result + lUIA[i + 1] + ':' + #9 + ws + #13#10;
									end
									else
									begin
                  	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[i + 1]  + ': ' + ws + ';</li>';
									end;
								end
								else
									if not HTMLout then ND.Value2 := None;
							except
								on E: Exception do
									// UIAs[31] := E.Message;
							end;
						end
						else if i = 53 then
						begin
							// node := nodes.AddChild(rNode, lUIA[54]);
              if not HTMLout then
              begin
								Node := SetNodeData(rNode, lUIA[54], '', nil, 0);
								Result := Result + lUIA[54];
              end
              else
              begin
              	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[54] + ';' + #13#10#9#9#9 + tab + '<ul>';
              end;
							try
								for t := 0 to 5 do
								begin
                	if not HTMLout then
                  begin
										SetNodeData(Node, lUIA[55 + t], UIAs[53 + t], nil, 0);
										Result := Result + #9 + lUIA[55 + t] + ':' +
											UIAs[53 + t] + #13#10;
                  end
                  else
                  begin
                  	Result := Result + #13#10#9#9#9 + tab + '<li>' +  lUIA[55 + t] + ': ' + UIAs[53 + t] + ';</li>';
                  end;
								end;
								if not HTMLout then TreeList1.Expanded[Node] := True
                else Result := Result + #13#10#9#9#9 + tab + '</ul>' +  #13#10#9#9 + tab + '</li>';
							except
								on E: Exception do
									// UIAs[31] := E.Message;
							end;
						end
						else if i = 54 then
						begin
            	if not HTMLout then
              begin
								SetNodeData(rNode, lUIA[61], UIAs[59], nil, 0);
								Result := Result + lUIA[61] + ':' + #9 + UIAs[59] + #13#10;
              end
              else
              begin
              	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[61] + ': ' + UIAs[59] + ';</li>';
              end;
						end
						else
						begin
            	if not HTMLout then
              begin
								SetNodeData(rNode, lUIA[i + 1], UIAs[i], nil, 0);
								Result := Result + lUIA[i + 1] + ':' + #9 + UIAs[i] + #13#10;
              end
              else
              begin
              	Result := Result + #13#10#9#9 + tab + '<li>' +  lUIA[i + 1] + ': ' + UIAs[i] + ';</li>';
              end;
						end;
					end;
				end;
			end;
			if not HTMLout then TreeList1.Expanded[rNode] := True;
		end;
  finally
  end;

end;


procedure TwndMSAAV.SetBalloonPos(X, Y:integer);
begin
    SendMessage(hWndTip, TTM_TRACKPOSITION,0, MAKELONG(X, Y));
    sendMessage(hWndTip, TTM_TRACKACTIVATE, 1, Integer(@ti));
end;

procedure TwndMSAAV.ShowBalloonTip(Control: TWinControl; Icon: integer; Title: string; Text: string; RC: TRect; X, Y:integer; Track:boolean = false);
var

  hhWnd: THandle;
  TP1: TPoint;
  Wnd: hwnd;
begin
    WindowFromAccessibleObject(iAcc, Wnd);
  hhWnd    := Control.Handle;
  hWndTip := CreateWindow(TOOLTIPS_CLASS, nil,
    WS_POPUP or TTS_NOPREFIX {or TTS_BALLOON} or TTS_ALWAYSTIP ,//or TTS_CLOSE ,
    RC.Left, RC.Top, 0, 0, hhWnd, 0, HInstance, nil);
  if hWndTip <> 0 then
  begin


    ti.cbSize := SizeOf(ti);
    if not Track then
        ti.uFlags := {TTF_CENTERTIP or }{TTF_IDISHWND  or }TTF_TRANSPARENT or TTF_SUBCLASS {or TTF_TRACK or TTF_ABSOLUTE}
    else
        ti.uFlags := {TTF_CENTERTIP or }{TTF_IDISHWND  or }TTF_TRANSPARENT or TTF_SUBCLASS or TTF_TRACK or TTF_ABSOLUTE;
    ti.hwnd := wnd{hhWnd};
    ti.uId := 1;
    ti.lpszText := PChar(Text);
    ti.Rect := RC;

    SendMessageW(hWndTip,TTM_SETMAXTIPWIDTH,0,640);
    SendMessage(hWndTip, TTM_ADDTOOL, 1, Integer(@ti));
    SendMessage(hWndTip, TTM_SETTITLE, Icon,  Integer(PChar(title)));

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
  tRC: TagRect;
    RC: TRect;
    vChild: variant;
    Mon : TMonitor;
    s, src, c, inner, outer: string;
    sLeft, sTop, i, iRes, iCnt: integer;
    iSP: IServiceProvider;
    iEle: IHTMLElement;
    isd: ISimpleDOMNode;
    PC: pchar;
    PC2:PChar;
    SI: Smallint;
    PU, PU2: PUINT;
    WD: Word;
    iText: ISimpleDOMText;
    tinfo: TToolInfo;
    wnd: HWND;
    pWnd: Pointer;
    monEx: TMonitorInfoEx;
    hm: HMonitor;
    iVW: IUIAutomationTreeWalker;
    pEle, cEle: IUIAutomationElement;
    hr: HResult;
begin
    if not Assigned(iAcc) and (not mnutvUIA.Checked) then Exit;
    if (mnutvUIA.Checked) then Exit;
    try



        if mnublnMSAA.Checked then
            s := s + sMSAATxt;//MSAAText4Tip;
        if mnublnIA2.Checked  and (not mnutvUIA.Checked) then
            s := s + SetIA2Text(iacc, false);
        if mnublnCode.Checked and (not mnutvUIA.Checked) then
					begin
						c := sHTML;
						inner := sTypeFF;
						outer := sTypeIE;
						hr := iAcc.QueryInterface(IID_IServiceProvider, iSP);
						if (hr = 0) and (Assigned(iSP)) then
						begin
							hr := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
							if (hr = 0) and (Assigned(iEle)) then
							begin

								if (SUCCEEDED(iSP.QueryService(IID_IHTMLElement,
									IID_IHTMLElement, iEle))) then
								begin
									sRC := iEle.outerHTML;
									sRC := copy(sRC, 0, ShowSrcLen);
									s := s + c + outer + ':' + #13#10 + sRC;
								end;
							end
							else
							begin
              	hr := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isd);
								if (hr = 0) and (Assigned(isd)) then
								begin
									isd.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);

									if WD <> 3 then
									begin
										PC := '';
										isd.get_innerHTML(PC);
										sRC := copy(PC, 0, ShowSrcLen);
										s := s + c + inner + #13#10 + sRC;
									end
									else
									begin
										iSP := isd as IServiceProvider;
										if SUCCEEDED(iSP.QueryInterface(IID_ISIMPLEDOMTEXT, iText))
										then
										begin
											iText.get_domText(PC);
											sRC := copy(PC, 0, ShowSrcLen);
											s := s + c + '(' + sTxt + ')' + #13#10 + sRC;
										end;
									end;
								end;
							end;
						end;
						// s := s + Src;
					end;

        if not mnutvUIA.Checked then
        begin
          vChild := varParent;//CHILDID_SELF;
          iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);

        end
        else
        begin
          if SUCCEEDED(UIEle.Get_CurrentBoundingRectangle(tRC)) then
          begin

            RC := Rect(tRc.left, tRC.top, tRC.right, tRC.bottom);
          end;
        end;
          sLeft := RC.Left;
          sTop := RC.Bottom;

        if hWndTip <> 0 then
        begin
            DestroyWindow(hWndTip);
        end;
            ShowBalloonTip(self, 1, 'Aviewer', s, rc, sLeft, sTop, True);
            SetBalloonPos(sLeft, sTop);
            if not mnutvUIA.Checked then
              WindowFromAccessibleObject(iAcc, Wnd)
            else
            begin
              UIAuto.Get_ControlViewWalker(iVW);
              if Assigned(iVW) then
              begin
                uiEle.Get_CurrentNativeWindowHandle(pWnd);
                cEle := UIEle;
                Wnd := HWND(pWnd);
                while Wnd = 0 do
                begin
                  iVW.GetParentElement(cELe, pEle);
                  if Assigned(pEle) then
                  begin
                    pEle.Get_CurrentNativeWindowHandle(pWnd);
                    Wnd := HWND(pWnd);
                    cEle := pEle;
                  end;
                end;
              end;
            end;
            tinfo.cbSize := SizeOf(tinfo);
            tinfo.hwnd := Wnd;
            tinfo.uId := 1;
            iRes := SendMessage(hWndTip, TTM_GETBUBBLESIZE , 0, Integer(@tinfo));

            if iRes > 0 then
            begin
                i := HIWORD(iRes);

                FillChar(monEx, SizeOf(TMonitorInfoEx), #0);
                monEx.cbSize := SizeOf(monEx);
                for iCnt := 0 to Screen.MonitorCount - 1 do
                begin

                  GetMonitorInfo(Screen.Monitors[iCnt].Handle, @monEx);
                  hm := MonitorFromWindow(hWndTip, MONITOR_DEFAULTTONEAREST);
                  // if PtInRect(monEx.rcMonitor , TP) then
                  if hm = Screen.Monitors[iCnt].Handle then
                  begin
                    Mon := Screen.Monitors[iCnt];//Screen.MonitorFromRect(rc);

                    if ((RC.Top + i + 20) > Mon.WorkareaRect.Bottom) then
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

                      SetBalloonPos(sLeft, sTop);
                      break;
                  end;
                end;
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
begin

    try
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        WndFocus.Shape1.Brush.Color := bkClr;

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



procedure TwndMSAAV.FormClose(Sender: TObject; var Action: TCloseAction);
var
    ini: TMemINiFile;
begin
	// UIAuto.Free;

	UIAuto := nil;

	NotifyWinEvent(EVENT_OBJECT_DESTROY, Handle, OBJID_CLIENT, 0);

	if hHook <> 0 then
	begin
		if UnhookWinEvent(hHook) then
		begin
			// FreeHookInstance(HookProc);
			hHook := 0;
		end;
	end;

	try
		if FileExists(SPath) then
		begin
			ini := TMemInifile.Create(SPath, TEncoding.Unicode);
			try
				ini.WriteInteger('Settings', 'Width', DoubleToInt(Width / ScaleX));
				ini.WriteInteger('Settings', 'Height', DoubleToInt(Height / ScaleY));
				ini.WriteInteger('Settings', 'Top', top);
				ini.WriteInteger('Settings', 'Left', left);
				// P2W, P4H: integer;
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
		end;
		if Assigned(TreeTH) then
		begin
			TreeTH.Terminate;
			TreeTH.WaitFor;
			TreeTH.Free;
			TreeTH := nil;
		end;
    if Assigned(UIATH) then
		begin
			UIATH.Terminate;
			UIATH.WaitFor;
			UIATH.Free;
			UIATH := nil;
		end;

    if Assigned(HTMLTH) then
		begin
			HTMLTH.Terminate;
			HTMLTH.WaitFor;
			HTMLTH.Free;
			HTMLTH := nil;
		end;

    if Assigned(SCList) then
    	SCList.Free;

		ClsNames.Free;
		TBList.Free;
    uTBList.Free;
		LangList.Free;
		CoUnInitialize;
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
	scList := nil;
		bFirstTime := True;
    bPFunc := False;
		dEventTime := 0;
    SystemCanSupportPerMonitorDpi(true);
    GetDCap(handle, Defx, Defy);
    GetWindowScale(Handle, DefX, DefY, ScaleX, ScaleY);
    cDPI := DoubleToInt(DefY * ScaleY);
    TreeList1.NodeDataSize := SizeOf(TNodeData);
    TreeTH := nil;
    APPDir :=  IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
    TransDir := IncludeTrailingPathDelimiter(AppDir + 'Languages');
    if not DirectoryExists(TransDir) then
      TransDir := IncludeTrailingPathDelimiter(AppDir + 'Lang');
    PageControl1.TabIndex := 0;
    mnuLang.Visible := FileExists(TransDir + 'Default.ini');
    if mnuLang.Visible then
        Transpath := TransDir + 'Default.ini'
    else
        Transpath := APPDir + ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
    DllPath := APPDir + 'IAccessible2Proxy.dll';
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
    TBList := TIntegerList.Create;
    uTBList := TIntegerList.Create;
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
    bSelMode := False;
    mnuSelMode.Checked := bSelMode;
    if mnuLang.Visible then
		begin
			if (FindFirst(TransDir + '*.ini', faAnyFile, Rec) = 0) then
			begin
				repeat
					if ((Rec.Name <> '.') and (Rec.Name <> '..')) then
					begin
						if ((Rec.Attr and faDirectory) = 0) then
						begin
							LangList.Add(Rec.Name);
						end;
					end;
				until (FindNext(Rec) <> 0);
			end;
		end;

    Load;
    ExecCmdLine;

    SizeChange;
    cDPI := DoubleToInt(DefY * ScaleY);

    bFirstTime := False;

end;


procedure TwndMSAAV.FormDestroy(Sender: TObject);
begin
    if Assigned(TreeTH) then TreeTH.Free;

end;

procedure TwndMSAAV.FormResize(Sender: TObject);
begin

    P2W := Panel2.Width;
    P4H := Panel4.Height;
end;


procedure TwndMSAAV.TreeList1Change(Sender: TObject; Node: TTreeNode);
var
	RC: TRect;
	iSP: IServiceProvider;
	Ele: IHTMLElement;
	SDom: ISImpleDOMNode;
begin

	if TreeList1.SelectedCount > 0 then
	begin

		if Node.Data <> nil then
		begin
			if acRect.Checked then
			begin
				if Assigned(CEle) then
				begin
					iSP := iAcc as IServiceProvider;
					if (SUCCEEDED(iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement,
						Ele))) then
						Ele.scrollIntoView(True);
				end
				else if Assigned(SDom) then
				begin
					iSP := iAcc as IServiceProvider;
					if (SUCCEEDED(iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE,
						SDom))) then
						SDom.scrollTo(True);
				end;
				Sleep(50);
				TTreeData(Node.Data^).Acc.accLocation(RC.left, RC.top, RC.right,
					RC.bottom, 0);
				ShowTargWnd(clBlue, RC);
			end;
		end;

	end;
end;

procedure TwndMSAAV.TreeList1Click(Sender: TObject);
begin
	if TreeList1.SelectedCount > 0 then
	begin
		if Assigned(UIAuto) and Assigned(iAcc) then
		begin
			if not not Assigned(iAcc) then
			begin
				if iDefIndex = integer(TreeList1.FocusedNode.Index) then
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

procedure TwndMSAAV.TreeList1FreeNode(Sender: TBaseVirtualTree;
  Node: PVirtualNode);
var
	ND: PNodeData;
begin
  ND := Sender.GetNodeData(Node);
  Finalize(ND^);
end;

procedure TwndMSAAV.TreeList1GetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
var
  ND: PNodeData;
begin
  if Sender.GetNodeData(Node) = nil then
    exit;
  if TextType = ttNormal then
  begin
    case Column of
      0:
      begin
        ND := Sender.GetNodeData(Node);
        CellText := ND.Value1;
      end;
      1:
      begin
        ND := Sender.GetNodeData(Node);
        CellText := ND.Value2;
      end;
    end;
  end;

end;

procedure TwndMSAAV.TreeList1KeyDown(Sender: TObject; var Key: Word;
	Shift: TShiftState);

begin
	if (TreeList1.SelectedCount > 0) and (Key = VK_RETURN) then
	begin
		if Assigned(UIAuto) and Assigned(iAcc) then
		begin
			if not not Assigned(iAcc) then
			begin

				if iDefIndex = integer(TreeList1.FocusedNode.Index) then
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

procedure TwndMSAAV.SetTreeMode(pNode: TTreeNode);
var
    b: boolean;
    RC: Tagrect;
begin

    if Assigned(TTreeData(pNode.Data^).Acc) then
    begin

      if pageControl1.ActivePageIndex = 0 then
      begin
      	TreeMode := True;
      	VarParent := TTreeData(pNode.Data^).iID;
      	iAcc :=  TTreeData(pNode.Data^).Acc;

          b := ShowMSAAText;
          GetNaviState;
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
          end
      end;
    end
    else if Assigned(TTreeData(pNode.Data^).uiEle) then
    begin
    	if pageControl1.ActivePageIndex = 1 then
      begin
      TreeMode := True;
      UIEle := TTreeData(pNode.Data^).uiEle;

      ShowText4UIA;
      GetNaviState;
      if (acRect.Checked) then
      begin
        UIEle.Get_CurrentBoundingRectangle(RC);
        ShowRectWnd2(clRed, Rect(rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top));
      end;
      if (acShowTip.Checked) then
      begin
        ShowTipWnd;
      end;
      end;
    end;

end;

procedure TwndMSAAV.TreeView1Expanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
{var
    Role: string;
    ovChild, ovRole: OleVariant;
    ws: widestring;
    i: integer;  }
begin
	{try
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
    end;  }
end;

procedure TwndMSAAV.TreeView1Addition(Sender: TObject; Node: TTreeNode);
var
	Role: string;
	ovChild, ovRole: OleVariant;
	ws: WideString;
	PC: PChar;
begin
	if not Assigned(Node.Data) then
		Exit;
	mnuTVSAll.Enabled := True;
	mnuTVOAll.Enabled := True;

  Node.ImageIndex := 7;
  Node.ExpandedImageIndex := 7;
  Node.SelectedIndex := 7;

  TTreeData(Node.Data^).Acc.Get_accName(TTreeData(Node.Data^).iID, ws);
		if ws = '' then
			ws := None;

		Role := Get_RoleText(TTreeData(Node.Data^).Acc, TTreeData(Node.Data^).iID);

		ovChild := TTreeData(Node.Data^).iID;
		ws := StringReplace(ws, #13, ' ', [rfReplaceAll]);
		ws := StringReplace(ws, #10, ' ', [rfReplaceAll]);
		Node.Text := ws + ' - ' + Role;
    {if (mnuAll.Checked) and (NodeTxt <> '') and (NodeTxt = Node.Text) and (not nSelected) then
    begin
    	Node.Expanded := True;
  		Node.Selected := True;
      nSelected := true;
    end;   }
		if SUCCEEDED(TTreeData(Node.Data^).Acc.Get_accRole(ovChild, ovRole)) then
		begin
			if VarHaveValue(ovRole) then
			begin
				if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
				begin
					Node.ImageIndex := TVarData(ovRole).VInteger - 1;
					Node.ExpandedImageIndex := Node.ImageIndex;
					Node.SelectedIndex := Node.ImageIndex;
					// showmessage(inttostr(rNode.ImageIndex));
				end;
			end;
		end;



end;

procedure TwndMSAAV.TreeView1Change(Sender: TObject; Node: TTreeNode);
begin

	if TreeView1.SelectionCount = 0 then
	begin
		mnuTVSSel.Enabled := false;
		mnuTVOSel.Enabled := false;
	end
	else
	begin
		mnuTVSSel.Enabled := True;
		mnuTVOSel.Enabled := True;
	end;
	if bSelMode then
		Exit;

	try
		if (((not Assigned(TreeTH)) or (TreeTH.Finished)) and (not mnutvUIA.Checked)) or
			(((not Assigned(UIATH)) or (UIATH.Finished)) and (mnutvUIA.Checked)) then
		begin
			SetTreeMode(Node);
		end;
	finally
	end;
end;

procedure TwndMSAAV.TreeView1Deletion(Sender: TObject; Node: TTreeNode);
begin
    Dispose(Node.Data);
end;

procedure TwndMSAAV.tbUIAAddition(Sender: TObject; Node: TTreeNode);
var
	Role: string;
	ovChild, ovRole: OleVariant;
	ws: WideString;
	PC: PChar;
begin

	if not Assigned(Node.Data) then
		Exit;
	mnuTVSAll.Enabled := True;
	mnuTVOAll.Enabled := True;

  Node.ImageIndex := 7;
  Node.ExpandedImageIndex := 7;
  Node.SelectedIndex := 7;

  TTreeData(Node.Data^).uiEle.Get_CurrentName(ws);
		if ws = '' then
			ws := None;
    ws := StringReplace(ws, #13, ' ', [rfReplaceAll]);
		ws := StringReplace(ws, #10, ' ', [rfReplaceAll]);
    Node.Text := ws;
		if SUCCEEDED(TTreeData(Node.Data^).uiEle.GetCurrentPropertyValue
			(UIA_LegacyIAccessibleRolePropertyId, ovRole)) then
		begin
			if VarHaveValue(ovRole) then
			begin
				if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
				begin
					Node.ImageIndex := TVarData(ovRole).VInteger - 1;
					Node.ExpandedImageIndex := Node.ImageIndex;
					Node.SelectedIndex := Node.ImageIndex;
					PC := StrAlloc(255);
					GetRoleTextW(ovRole, PC, StrBufSize(PC));
					Node.Text := ws + ' - ' + PC;
					StrDispose(PC);
				end;
			end;
		end;
end;

procedure TwndMSAAV.tbUIAChange(Sender: TObject; Node: TTreeNode);
begin
	if tbUIA.SelectionCount = 0 then
	begin
		mnuTVSSel.Enabled := false;
		mnuTVOSel.Enabled := false;
	end
	else
	begin
		mnuTVSSel.Enabled := True;
		mnuTVOSel.Enabled := True;
	end;
	if bSelMode then
		Exit;

	try
		if (((not Assigned(TreeTH)) or (TreeTH.Finished)) and (not mnutvUIA.Checked)) or
			(((not Assigned(UIATH)) or (UIATH.Finished)) and (mnutvUIA.Checked)) then
		begin
			SetTreeMode(Node);

		end;
	finally
	end;
end;

procedure TwndMSAAV.tbUIADeletion(Sender: TObject; Node: TTreeNode);
begin
	Dispose(Node.Data);
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
begin
	Clipboard.AsText := sMSAATxt + #10#13 + sARIATxt + #13#10 + sHTMLTxt + #10#13 + sIA2txt + #10#13 + sUIATxt;
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
	if HelpURL <> '' then
	begin
		ShellExecuteW(Handle, 'open', pWidechar(HelpURL), nil, nil, SW_SHOW);
	end;
end;



procedure TwndMSAAV.acListFocusExecute(Sender: TObject);
begin
  TreeList1.SetFocus;
end;

procedure TwndMSAAV.acMSAAModeExecute(Sender: TObject);
begin
	mnutvMSAA.Checked := True;
  mnutvUIA.Checked := False;
  mnutvBoth.Checked := False;

  TabSheet1.TabVisible := true;
  TabSheet2.TabVisible := False;
  if Assigned(TreeTH) then
	begin
		TreeTH.Terminate;
		TreeTH.WaitFor;
		TreeTH.Free;
		TreeTH := nil;
	end;

	if Assigned(WndFocus) then
	begin
		WndFocus.Visible := false;
	end;
	if Assigned(WndTip) then
	begin
		WndTip.Visible := false;
	end;
end;

procedure TwndMSAAV.mnutvUIAClick(Sender: TObject);
begin

	mnutvMSAA.Checked := False;
  mnutvUIA.Checked := False;
  mnutvBoth.Checked := False;
  TabSheet1.TabVisible := true;
  TabSheet2.TabVisible := true;
	if Sender = mnutvMSAA then
  begin
  	mnutvMSAA.Checked := True;
    TabSheet2.TabVisible := False;
  end
  else if Sender = mnutvUIA then
  begin
  	mnutvUIA.Checked := True;
    TabSheet1.TabVisible := False;
  end
  else
  begin
    mnutvBoth.Checked := True;
    Pagecontrol1.ActivePageIndex := 0;
  end;

	if Assigned(TreeTH) then
	begin
		TreeTH.Terminate;
		TreeTH.WaitFor;
		TreeTH.Free;
		TreeTH := nil;
	end;

	if Assigned(WndFocus) then
	begin
		WndFocus.Visible := false;
	end;
	if Assigned(WndTip) then
	begin
		WndTip.Visible := false;
	end;

end;

procedure TwndMSAAV.ExecMSAAMode(Sender: TObject);
var
	pAcc: IAccessible;
	s: string;
	tc: TComponent;
	iDis: iDispatch;
	iRes, iDir: integer;
	ov: OleVariant;
	sNode: TTreeNode;
  tTB: TAccTreeView;
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
  if PageControl1.ActivePageIndex = 0 then
  	tTB := TreeView1
  else
  	tTB := tbUIA;
	if (s = 'tbparent') or (s = 'acparent') then
	begin
		try
			if tTB.Items.Count > 0 then
			begin
				if (tTB.SelectionCount > 0) then
				begin
					if tTB.Selected.Parent <> nil then
					begin
						sNode := tTB.Selected.Parent;
					end;
				end;
			end;
		except
			on E: Exception do
				ShowErr(E.Message);
		end;
	end
	else if (s = 'tbchild') or (s = 'acchild') then
	begin
		if tTB.Items.Count > 0 then
		begin
			if (tTB.SelectionCount > 0) then
			begin
				if tTB.Selected.HasChildren then
				begin
					sNode := tTB.Selected.getFirstChild;

				end;
			end;
		end;
	end
	else if (s = 'tbprevs') or (s = 'acprevs') or (s = 'tbnexts') or
		(s = 'acnexts') then
	begin
		if tTB.Items.Count > 0 then
		begin
			if (tTB.SelectionCount > 0) then
			begin
				if (s = 'tbprevs') or (s = 'acprevs') then
				begin
					if tTB.Selected.getPrevSibling <> nil then
					begin
						sNode := tTB.Selected.getPrevSibling;

					end;
				end
				else
				begin
					if tTB.Selected.getNextSibling <> nil then
					begin
						sNode := tTB.Selected.getNextSibling;

					end;
				end;
			end;
		end;
	end
	else
		pAcc := nil;
	try
		if Assigned(sNode) then
		begin
			Treemode := True;
			sNode.Selected := True;
		end;
	except
		on E: Exception do
		begin
			ShowErr(E.Message);
		end;
	end;
end;

procedure TwndMSAAV.acNextSExecute(Sender: TObject);
begin
    TreeList1.Clear;
    Memo1.ClearAll;
    ExecMSAAMode(acNextS);
end;

procedure TwndMSAAV.acPrevSExecute(Sender: TObject);
begin
    TreeList1.Clear;
    Memo1.ClearAll;
    ExecMSAAMode(acPrevS);
end;

procedure TwndMSAAV.acOnlyFocusExecute(Sender: TObject);
begin
    acOnlyFOcus.Checked := not acOnlyFOcus.Checked;
    if acOnlyFocus.Checked then
        mnutvUIA.Checked := false;
    ExecOnlyFocus;
end;

procedure TwndMSAAV.acChildExecute(Sender: TObject);
begin
    TreeList1.Clear;
    Memo1.ClearAll;
    ExecMSAAMode(acChild);
end;


procedure TwndMSAAV.acParentExecute(Sender: TObject);
begin
    TreeList1.Clear;
    Memo1.ClearAll;
    ExecMSAAMode(acParent);
end;



procedure TwndMSAAV.acRectExecute(Sender: TObject);
var
  RC: TagRect;
begin
     acRect.Checked := not acRect.Checked;
    if acRect.Checked then
    begin
        if not Assigned(WndFocus) then
            WndFocus := TWndFocusRect.Create(self);
        if not mnutvUIA.Checked then
          ShowRectWnd(clRed)
        else
        begin
          UIEle.Get_CurrentBoundingRectangle(RC);
          ShowRectWnd2(clRed, Rect(rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top));
        end;
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
    d, lbtb, lbmf, mfTV, mfLV, mfST, mf3ctrl: string;
    sc: word;
    ini: TMemINiFile;
    SCD: PSCData;
    cNode, rNode: PVirtualNode;
begin
  FormStyle := fsNormal;
  timer1.Enabled := false;
    WndSet := TWndSet.Create(self);
    try
        WndSet.Load_Str(TransPath);
        WndSet.Font := Font;
        WndSet.DefFont := DefFont;
        {WndSet.DefY := DefY;
        WndSet.DefX := DefX;
        WndSet.ScaleX := ScaleX;
        WndSet.ScaleY := ScaleY;  }
        WndSet.DefW := WndSet.Width;
        WndSet.DefH := WndSet.Height;
        WndSet.FontDialog1.Font := Font;
        {WndSet.cDPI := DoubleToInt(DefY * ScaleY);
        WndSet.sDPI := WndSet.cDPI; }



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
        WndSet.ChkList.Items.Add(lIA2[14]);
        b := Trunc(IntPower(2, 10));
        if flgIA2 and b <> 0 then
        begin
          WndSet.chkList.Checked[23] := True;
        end;

        WndSet.ChkList.Items.Add(lIA2[15]);
        b := Trunc(IntPower(2, 11));
        if flgIA2 and b <> 0 then
        begin
          WndSet.chkList.Checked[24] := True;
        end;


        WndSet.ChkList.Items.Add(lUIA[0]);
        WndSet.ChkList.Header[WndSet.ChkList.Items.Count - 1] := True;
        for i := 1 to 55 do
        begin
          if i = 55 then
            WndSet.ChkList.Items.Add(lUIA[61])
          else
            WndSet.ChkList.Items.Add(lUIA[i]);

            if i < 32 then
            begin
                b := Trunc(IntPower(2, i-1));
                if flgUIA and b <> 0 then
                begin
                    WndSet.chkList.Checked[i + 25] := True;
                end;
            end
            else
            begin
                b := Trunc(IntPower(2, i-32));
                if flgUIA2 and b <> 0 then
                begin
                    WndSet.chkList.Checked[i + 25] := True;
                end;
            end;
        end;
        WndSet.cmbShortCut.Items.Add(None);
        for i := 0 to 25 do
        begin
        	WndSet.cmbShortCut.Items.Add('Ctrl+' + Chr(65 + i));
        end;

        for i := 0 to 25 do
        begin
        	WndSet.cmbShortCut.Items.Add('Ctrl+Alt+' + Chr(65 + i));
        end;

        for i := 1 to 12 do
        begin
        	WndSet.cmbShortCut.Items.Add('F' + inttostr(i));
        end;

        for i := 1 to 12 do
        begin
        	WndSet.cmbShortCut.Items.Add('Ctrl+F' + inttostr(i));
        end;

        for i := 1 to 12 do
        begin
        	WndSet.cmbShortCut.Items.Add('Ctrl+Shift+F' + inttostr(i));
        end;

        for i := 1 to 12 do
        begin
        	WndSet.cmbShortCut.Items.Add('Shift+F' + inttostr(i));
        end;

        for i := 1 to 12 do
        begin
        	WndSet.cmbShortCut.Items.Add('Alt+F' + inttostr(i));
        end;

        ini := TMemIniFile.Create(TransPath, TEncoding.UTF8);
        try
        	lbmf := ini.ReadString('SetDLG', 'lbmf', 'Move focus to');
          lbtb := ini.ReadString('SetDLG', 'lbtb', 'Toolbar functions');

          mfTV := ini.ReadString('SetDLG', 'mfTreeView', 'TreeView');
        	//WndSet.clShortcut.Checked[WndSet.clShortcut.Items.Count - 1] := (acTreeFocus.ShortCut <> 0);


          mfLV := ini.ReadString('SetDLG', 'mfListView', 'ListView');
          //WndSet.clShortcut.Checked[WndSet.clShortcut.Items.Count - 1] := (acListFocus.ShortCut <> 0);


          mfST := ini.ReadString('SetDLG', 'mfTextBox', 'Source Text');
          //WndSet.clShortcut.Checked[WndSet.clShortcut.Items.Count - 1] := (acMemoFocus.ShortCut <> 0);

          mf3ctrl := ini.ReadString('SetDLG', 'mf3ctrls', 'between 3 controls');
          //WndSet.clShortcut.Checked[WndSet.clShortcut.Items.Count - 1] := (acTriFocus.ShortCut <> 0);


				finally
					ini.Free;
				end;
        WndSet.clShortcut.NodeDataSize := SizeOf(TSCData);

        rNode := WndSet.clShortcut.AddChild(nil);
        SCD := rNode.GetData;
        SCD^.Name := lbtb;
  			SCD^.SCKey := 0;
        WndSet.TFNode := rNode;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbFocus.Hint;
  			SCD^.SCKey := acFocus.ShortCut;
        SCD^.actName := acFocus.Name;
        WndSet.clShortcut.Selected[cNode] := True;
        WndSet.clShortcut.OnClick(WndSet.clShortcut);

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbCursor.Hint;
  			SCD^.SCKey := acCursor.ShortCut;
        SCD^.actName := acCursor.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbRectAngle.Hint;
  			SCD^.SCKey := acRect.ShortCut;
        SCD^.actName := acRect.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbShowtip.Hint;
  			SCD^.SCKey := acShowtip.ShortCut;
        SCD^.actName := acShowtip.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbCopy.Hint;
  			SCD^.SCKey := acCopy.ShortCut;
        SCD^.actName := acCopy.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbOnlyFocus.Hint;
  			SCD^.SCKey := acOnlyFocus.ShortCut;
        SCD^.actName := acOnlyFocus.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbParent.Hint;
  			SCD^.SCKey := acParent.ShortCut;
        SCD^.actName := acParent.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbChild.Hint;
  			SCD^.SCKey := acChild.ShortCut;
        SCD^.actName := acChild.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbPrevS.Hint;
  			SCD^.SCKey := acPrevS.ShortCut;
        SCD^.actName := acPrevS.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbNextS.Hint;
  			SCD^.SCKey := acNextS.ShortCut;
        SCD^.actName := acNextS.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbHelp.Hint;
  			SCD^.SCKey := acHelp.ShortCut;
        SCD^.actName := acHelp.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbMSAAMode.Hint;
  			SCD^.SCKey := acMSAAMode.ShortCut;
        SCD^.actName := acMSAAMode.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := tbRegister.Hint;
  			SCD^.SCKey := acSetting.ShortCut;
        SCD^.actName := acSetting.Name;
        WndSet.clShortcut.Expanded[rNode] := True;

        rNode := WndSet.clShortcut.AddChild(nil);
        SCD := rNode.GetData;
        SCD.Name := lbmf;
  			SCD.SCKey := 0;
        WndSet.MFNode := rNode;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := mfTV;
  			SCD^.SCKey := acTreeFocus.ShortCut;
        SCD^.actName := acTreeFocus.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := mfLV;
  			SCD^.SCKey := acListFocus.ShortCut;
        SCD^.actName := acListFocus.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := mfST;
  			SCD^.SCKey := acMemoFocus.ShortCut;
        SCD^.actName := acMemoFocus.Name;

        cNode := WndSet.clShortcut.AddChild(rNode);
        SCD := cNode.GetData;
  			SCD^.Name := mf3ctrl;
  			SCD^.SCKey := ac3ctrls.ShortCut;
        SCD^.actName := ac3ctrls.Name;

        WndSet.clShortcut.Expanded[rNode] := True;
        WndSet.SizeChange;
        if WndSet.ShowModal = mrOK then
        begin
            Font := WndSet.Font;
            DefFont := WndSet.DefFont;
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
            for i := 0 to 11 do
            begin
                if WndSet.chkList.Checked[i + 13] then
                begin
                    b := Trunc(IntPower(2, i));
                    flgIA2 := flgIA2 or b;
                end;
            end;
            flgUIA := 0;
            flgUIA2 := 0;
            for i := 0 to 54 do
            begin
                if WndSet.chkList.Checked[i + 26] then
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
                ini.WriteInteger('Settings', 'FontSize', DefFont);
                Ini.WriteInteger('Settings', 'Charset', Font.Charset);
                ini.WriteInteger('Settings', 'flgMSAA', flgMSAA);
                ini.WriteInteger('Settings', 'flgIA2', flgIA2);
                ini.WriteInteger('Settings', 'flgUIA', flgUIA);
                ini.WriteInteger('Settings', 'flgUIA2', flgUIA2);
                ini.WriteBool('Settings', 'ExTooltip', Extip);

                rNode := WndSet.TFNode;
								cNode := rNode.FirstChild;
                SCD := WndSet.clShortcut.GetNodeData(cNode);
                ini.WriteInteger('Shortcut', SCD^.actName, SCD^.SCKey);
            		for i := 1 to rNode.ChildCount - 1 do
								begin
									cNode := cNode.NextSibling;
									if cNode = nil then
										break
									else
                  begin
                  	SCD := WndSet.clShortcut.GetNodeData(cNode);
                		ini.WriteInteger('Shortcut', SCD^.actName, SCD^.SCKey);
                  end;
								end;

								rNode := WndSet.MFNode;
								cNode := rNode.FirstChild;
                SCD := WndSet.clShortcut.GetNodeData(cNode);
                ini.WriteInteger('Shortcut', SCD^.actName, SCD^.SCKey);
								for i := 1 to rNode.ChildCount - 1 do
								begin
									cNode := cNode.NextSibling;
									if cNode = nil then
										break
									else
									begin
                  	SCD := WndSet.clShortcut.GetNodeData(cNode);
                		ini.WriteInteger('Shortcut', SCD^.actName, SCD^.SCKey);
                  end;
								end;

                ini.UpdateFile;

                acFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acFocus.Name, 115));
        				acCursor.ShortCut := Word(ini.ReadInteger('Shortcut', acCursor.Name, 116));
                acRect.ShortCut := Word(ini.ReadInteger('Shortcut', acRect.Name, 117));
        				acShowtip.ShortCut := Word(ini.ReadInteger('Shortcut', acShowtip.Name, 114));
        				acCopy.ShortCut := Word(ini.ReadInteger('Shortcut', acCopy.Name, 118));
        				acOnlyfocus.ShortCut := Word(ini.ReadInteger('Shortcut', acOnlyfocus.Name, 119));
        				acParent.ShortCut := Word(ini.ReadInteger('Shortcut', acParent.Name, 120));
        				acChild.ShortCut := Word(ini.ReadInteger('Shortcut', acChild.Name, 121));
        				acPrevS.ShortCut := Word(ini.ReadInteger('Shortcut', acPrevS.Name, 122));
        				acNextS.ShortCut := Word(ini.ReadInteger('Shortcut', acNextS.Name, 123));
        				acHelp.ShortCut := Word(ini.ReadInteger('Shortcut', acHelp.Name, 112));
        				acMSAAMode.ShortCut := Word(ini.ReadInteger('Shortcut', acMSAAMode.Name, 8308));
        				acSetting.ShortCut := Word(ini.ReadInteger('Shortcut', acSetting.Name, 8309));
        				acTreeFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acTreeFocus.Name, 8304));
        				acListFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acListFocus.Name, 8305));
        				acMemoFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acMemoFocus.Name, 8306));
        				ac3ctrls.ShortCut := Word(ini.ReadInteger('Shortcut', ac3ctrls.Name, 8307));

            finally
                ini.Free;
            end;
            SizeChange;
        end;
    finally
        WndSet.Free;
        FormStyle := fsStayOnTop;
        timer1.Enabled := true;
    end;
end;



procedure TwndMSAAV.acMemoFocusExecute(Sender: TObject);
begin
  Memo1.SetFocus;
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

procedure TwndMSAAV.acTreeFocusExecute(Sender: TObject);
begin
  TreeView1.SetFocus;
end;

procedure TwndMSAAV.ac3ctrlsExecute(Sender: TObject);
begin
  if TreeView1.Focused then
    TreeList1.SetFocus
  else if TreeList1.Focused then
    Memo1.SetFocus
  else if Memo1.Focused then
    TreeView1.SetFocus
  else
    TreeView1.SetFocus;
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
        if (TreeView1.Items.Count > 0) and (Pagecontrol1.ActivePageIndex = 0) then
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

        end;
        if (tbUIA.Items.Count > 0) and (Pagecontrol1.ActivePageIndex = 1) then
        begin
            if (tbUIA.SelectionCount > 0) then
            begin
                if tbUIA.Selected.Parent <> nil then
                begin
                    tbParent.Enabled := true;
                end;
                if tbUIA.Selected.HasChildren then
                begin
                    tbChild.Enabled := true;
                end;
                if tbUIA.Selected.getNextSibling <> nil then
                begin
                    tbNextS.Enabled := true;
                end;
                if tbUIA.Selected.getPrevSibling<> nil then
                begin
                    tbPrevS.Enabled := true;
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

function GetFileVersionString(FileName: string): String;
var
  dwHandle  : Cardinal;
  pInfo     : Pointer;
  InfoSize  : DWORD;
  pFileInfo : PVSFixedFileInfo;
  iVer      : array[0..2] of Cardinal;
begin
  Result := '';

  InfoSize := GetFileVersionInfoSize(PChar(FileName), dwHandle);
  if InfoSize = 0 then exit;

  GetMem(pInfo, InfoSize);
  try
    GetFileVersionInfo(PChar(FileName), 0, InfoSize, pInfo);
    VerQueryValue(pInfo, PathDelim, Pointer(pFileInfo), InfoSize);

    iVer[0] := pFileInfo.dwFileVersionMS shr 16;
    iVer[1] := pFileInfo.dwFileVersionMS and $FFFF;
    iVer[2] := pFileInfo.dwFileVersionLS shr 16;
    Result := Format('%d.%d.%d', [iVer[0], iVer[1], iVer[2]]);
  finally
    FreeMem(pInfo, InfoSize);
  end;
end;


function TwndMSAAV.LoadLang: string;
var
	  ini: TMemIniFile;
    d, Msg, UIA_fail: string;
begin
    Result := 'CoCreateInstance failed.';
    ini := TMemIniFile.Create(TransPath, TEncoding.UTF8);
    ClsNames := TStringList.Create;
	    try

            lMSAA[0] := ini.ReadString('MSAA', 'MSAA', 'MS Active Accessibility');
            lMSAA[1] := ini.ReadString('MSAA', 'accName','accName');
            lMSAA[2] := ini.ReadString('MSAA', 'accRole','accRole');
            lMSAA[3] := ini.ReadString('MSAA', 'accState','accState');
            lMSAA[4] := ini.ReadString('MSAA', 'accDescription','accDescription');
            lMSAA[5] := ini.ReadString('MSAA', 'accDefaultAction','accDefaultAction');
            lMSAA[6] := ini.ReadString('MSAA', 'accValue','accValue');
            lMSAA[7] := ini.ReadString('MSAA', 'accParent','accParent');
            lMSAA[8] := ini.ReadString('MSAA', 'accChildCount','accChildCount');
            lMSAA[9] := ini.ReadString('MSAA', 'accHelp','accHelp');
            lMSAA[10] := ini.ReadString('MSAA', 'accHelpTopic','accHelpTopic');
            lMSAA[11] := ini.ReadString('MSAA', 'accKeyboardShortcut','accKeyboardShortcut');


            lIA2[0] := ini.ReadString('IA2', 'IA2', 'IAccessible2');
            lIA2[1] := ini.ReadString('IA2', 'Name','Name');
            lIA2[2] := ini.ReadString('IA2', 'Role','Role');
            lIA2[3] := ini.ReadString('IA2', 'States','States');
            lIA2[4] := ini.ReadString('IA2', 'Description','Description');
            lIA2[5] := ini.ReadString('IA2', 'Relations','Relations');
            //lIA2[6] := LoadTranslation('IA2', 'RelationTargets','Relation Targets');
            lIA2[6] := ini.ReadString('IA2', 'Attributes','Object Attributes');
            lIA2[7] := ini.ReadString('IA2', 'Value','Value');
            lIA2[8] := ini.ReadString('IA2', 'LocalizedExtendedRole','LocalizedExtendedRole');
            lIA2[9] := ini.ReadString('IA2', 'LocalizedExtendedStates','LocalizedExtendedStates');
            lIA2[10] := ini.ReadString('IA2', 'RangeValue','RangeValue');
            lIA2[11] := ini.ReadString('IA2', 'RV_Value','Value');
            lIA2[12] := ini.ReadString('IA2', 'RV_Minimum','Minimum');
            lIA2[13] := ini.ReadString('IA2', 'RV_Maximum','Maximum');
            lIA2[14] := ini.ReadString('IA2', 'Textattributes','Textattributes');
            lIA2[15] := ini.ReadString('IA2', 'uniqueID','uniqueID');


            lUIA[0] := ini.ReadString('UIA', 'UIA', 'UIAutomation');
            lUIA[1] := ini.ReadString('UIA', 'CurrentAcceleratorKey','CurrentAcceleratorKey');
            lUIA[2] := ini.ReadString('UIA', 'CurrentAccessKey','CurrentAccessKey');
            lUIA[3] := ini.ReadString('UIA', 'CurrentAriaProperties','CurrentAriaProperties');
            lUIA[4] := ini.ReadString('UIA', 'CurrentAriaRole','CurrentAriaRole');
            lUIA[5] := ini.ReadString('UIA', 'CurrentAutomationId','CurrentAutomationId');
            lUIA[6] := ini.ReadString('UIA', 'CurrentBoundingRectangle','CurrentBoundingRectangle');
            lUIA[7] := ini.ReadString('UIA', 'CurrentClassName','CurrentClassName');
            lUIA[8] := ini.ReadString('UIA', 'CurrentControlType','CurrentControlType');
            lUIA[9] := ini.ReadString('UIA', 'CurrentCulture','CurrentCulture');
            lUIA[10] := ini.ReadString('UIA', 'CurrentFrameworkId','CurrentFrameworkId');
            lUIA[11] := ini.ReadString('UIA', 'CurrentHasKeyboardFocus','CurrentHasKeyboardFocus');
            lUIA[12] := ini.ReadString('UIA', 'CurrentHelpText','CurrentHelpText');
            lUIA[13] := ini.ReadString('UIA', 'CurrentIsControlElement','CurrentIsControlElement');
            lUIA[14] := ini.ReadString('UIA', 'CurrentIsContentElement','CurrentIsContentElement');
            lUIA[15] := ini.ReadString('UIA', 'CurrentIsDataValidForForm','CurrentIsDataValidForForm');
            lUIA[16] := ini.ReadString('UIA', 'CurrentIsEnabled','CurrentIsEnabled');
            lUIA[17] := ini.ReadString('UIA', 'CurrentIsKeyboardFocusable','CurrentIsKeyboardFocusable');
            lUIA[18] := ini.ReadString('UIA', 'CurrentIsOffscreen','CurrentIsOffscreen');
            lUIA[19] := ini.ReadString('UIA', 'CurrentIsPassword','CurrentIsPassword');
            lUIA[20] := ini.ReadString('UIA', 'CurrentIsRequiredForForm','CurrentIsRequiredForForm');
            lUIA[21] := ini.ReadString('UIA', 'CurrentItemStatus','CurrentItemStatus');
            lUIA[22] := ini.ReadString('UIA', 'CurrentItemType','CurrentItemType');
            lUIA[23] := ini.ReadString('UIA', 'CurrentLocalizedControlType','CurrentLocalizedControlType');
            lUIA[24] := ini.ReadString('UIA', 'CurrentName','CurrentName');
            lUIA[25] := ini.ReadString('UIA', 'CurrentNativeWindowHandle','CurrentNativeWindowHandle');
            lUIA[26] := ini.ReadString('UIA', 'CurrentOrientation','CurrentOrientation');
            lUIA[27] := ini.ReadString('UIA', 'CurrentProcessId','CurrentProcessId');
            lUIA[28] := ini.ReadString('UIA', 'CurrentProviderDescription','CurrentProviderDescription');
            lUIA[29] := ini.ReadString('UIA', 'CurrentControllerFor','CurrentControllerFor');
            lUIA[30] := ini.ReadString('UIA', 'CurrentDescribedBy','CurrentDescribedBy');
            lUIA[31] := ini.ReadString('UIA', 'CurrentFlowsTo','CurrentFlowsTo');
            lUIA[32] := ini.ReadString('UIA', 'CurrentLabeledBy','CurrentLabeledBy');
            lUIA[33] := ini.ReadString('UIA', 'CurrentLiveSetting','CurrentLiveSetting');
            //Added 2014/06/23
            lUIA[34] := ini.ReadString('UIA', 'IsDockPatternAvailable','IsDockPatternAvailable');
            lUIA[35] := ini.ReadString('UIA', 'IsExpandCollapsePatternAvailable','IsExpandCollapsePatternAvailable');
            lUIA[36] := ini.ReadString('UIA', 'IsGridItemPatternAvailable','IsGridItemPatternAvailable');
            lUIA[37] := ini.ReadString('UIA', 'IsGridPatternAvailable','IsGridPatternAvailable');
            lUIA[38] := ini.ReadString('UIA', 'IsInvokePatternAvailable','IsInvokePatternAvailable');
            lUIA[39] := ini.ReadString('UIA', 'IsMultipleViewPatternAvailable','IsMultipleViewPatternAvailable');
            lUIA[40] := ini.ReadString('UIA', 'IsRangeValuePatternAvailable','IsRangeValuePatternAvailable');
            lUIA[41] := ini.ReadString('UIA', 'IsScrollPatternAvailable','IsScrollPatternAvailable');
            lUIA[42] := ini.ReadString('UIA', 'IsScrollItemPatternAvailable','IsScrollItemPatternAvailable');
            lUIA[43] := ini.ReadString('UIA', 'IsSelectionItemPatternAvailable','IsSelectionItemPatternAvailable');
            lUIA[44] := ini.ReadString('UIA', 'IsSelectionPatternAvailable','IsSelectionPatternAvailable');
            lUIA[45] := ini.ReadString('UIA', 'IsTablePatternAvailable','IsTablePatternAvailable');
            lUIA[46] := ini.ReadString('UIA', 'IsTableItemPatternAvailable','IsTableItemPatternAvailable');
            lUIA[47] := ini.ReadString('UIA', 'IsTextPatternAvailable','IsTextPatternAvailable');
            lUIA[48] := ini.ReadString('UIA', 'IsTogglePatternAvailable','IsTogglePatternAvailable');
            lUIA[49] := ini.ReadString('UIA', 'IsTransformPatternAvailable','IsTransformPatternAvailable');
            lUIA[50] := ini.ReadString('UIA', 'IsValuePatternAvailable','IsValuePatternAvailable');
            lUIA[51] := ini.ReadString('UIA', 'IsWindowPatternAvailable','IsWindowPatternAvailable');
            lUIA[52] := ini.ReadString('UIA', 'IsItemContainerPatternAvailable','IsItemContainerPatternAvailable');
            lUIA[53] := ini.ReadString('UIA', 'IsVirtualizedItemPatternAvailable','IsVirtualizedItemPatternAvailable');
            //Added 2014/06/25
            lUIA[54] := ini.ReadString('UIA', 'RangeValue','RangeValue');
            lUIA[55] := ini.ReadString('UIA', 'RV_Value','Value');
            lUIA[56] := ini.ReadString('UIA', 'RV_IsReadOnly','IsReadOnly');
            lUIA[57] := ini.ReadString('UIA', 'RV_Minimum','Minimum');
            lUIA[58] := ini.ReadString('UIA', 'RV_Maximum','Maximum');
            lUIA[59] := ini.ReadString('UIA', 'RV_LargeChange','LargeChange');
            lUIA[60] := ini.ReadString('UIA', 'RV_SmallChange','SmallChange');
            lUIA[61] := ini.ReadString('UIA', 'StyleID','Style');

            StyleID[0] := ini.ReadString('UIA', 'StyleId_Custom','A custom style.');
            StyleID[1] := ini.ReadString('UIA', 'StyleId_Heading1','A first level heading.');
            StyleID[2] := ini.ReadString('UIA', 'StyleId_Heading2','A second level heading.');
            StyleID[3] := ini.ReadString('UIA', 'StyleId_Heading3','A third level heading.');
            StyleID[4] := ini.ReadString('UIA', 'StyleId_Heading4','A fourth level heading.');
            StyleID[5] := ini.ReadString('UIA', 'StyleId_Heading5','A fifth level heading.');
            StyleID[6] := ini.ReadString('UIA', 'StyleId_Heading6','A sixth level heading.');
            StyleID[7] := ini.ReadString('UIA', 'StyleId_Heading7','A seventh level heading.');
            StyleID[8] := ini.ReadString('UIA', 'StyleId_Heading8','An eighth level heading.');
            StyleID[9] := ini.ReadString('UIA', 'StyleId_Heading9','A ninth level heading.');
            StyleID[10] := ini.ReadString('UIA', 'StyleId_Title','A title.');
            StyleID[11] := ini.ReadString('UIA', 'StyleId_Subtitle','A subtitle.');
            StyleID[12] := ini.ReadString('UIA', 'StyleId_Normal','Normal style.');
            StyleID[13] := ini.ReadString('UIA', 'StyleId_Emphasis','Text that is emphasized.');
            StyleID[14] := ini.ReadString('UIA', 'StyleId_Quote','A quotation.');
            StyleID[15] := ini.ReadString('UIA', 'StyleId_BulletedList','A list with bulleted items.');
            StyleID[16] := ini.ReadString('UIA', 'StyleId_NumberedList','A list with numbered items.');

            HelpURL := ini.ReadString('General', 'Help_URL', 'http://www.google.com/');
            Caption := ini.ReadString('General', 'MSAA_Caption', 'Accessibility Viewer');
            Caption := Caption + ' - ' + GetFileVersionString(application.ExeName);
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
            DefFont := Font.Size;
            mnuSelD.Caption := ini.ReadString('General', 'MSAA_gbSelDisplay', 'Select Display');
            tbFocus.Hint := ini.ReadString('General', 'MSAA_tbFocusHint', 'Watch Focus');
            tbCursor.Hint := ini.ReadString('General', 'MSAA_tbCursorHint', 'Watch Cursor');
            tbRectAngle.Hint := ini.ReadString('General', 'MSAA_tbRectAngleHint', 'Show Highlight Rectangle');

            tbShowtip.Hint :=  ini.ReadString('General', 'MSAA_tbBalloonHint', 'Show Balloon tip');
            acShowTip.Hint := tbShowtip.Hint;
            tbCopy.Hint := ini.ReadString('General', 'MSAA_tbCopyHint', 'Copy Text to Clipborad');
            tbOnlyFocus.Hint := ini.ReadString('General', 'MSAA_tbFocusOnly', 'Focus rectangle only');

            tbParent.Hint := ini.ReadString('General', 'MSAA_tbParentHint', 'Navigates to parent object');
            tbChild.Hint := ini.ReadString('General', 'MSAA_tbChildHint', 'Navigates to first child object');
            tbPrevS.Hint := ini.ReadString('General', 'MSAA_tbPrevSHint', 'Navigates to previous sibling object');
            tbNextS.Hint := ini.ReadString('General', 'MSAA_tbNextSHint', 'Navigates to next sibling object');
            tbHelp.Hint := ini.ReadString('General', 'MSAA_tbHelpHint', 'Show online help');
            //tbHelp.Hint := tbHelp.Hint + '(' +HelpURL + ')';

            tbMSAAMode.Hint := ini.ReadString('General', 'MSAA_tbMSAAModeHint', 'UIA mode');
            tbRegister.Hint := ini.ReadString('General', 'MSAA_tbMSAASetHint', 'Show Setting Dialog');
            mnuLang.Caption := ini.ReadString('General', 'mnuLang', '&Language');
            mnuView.Caption := ini.ReadString('General', 'mnuView', '&View');
            mnuMSAA.Caption := ini.ReadString('General', 'MSAA_cbMSAA', 'MSAA');
            mnuARIA.Caption := ini.ReadString('General', 'MSAA_cbARIA', 'ARIA');
            mnuHTML.Caption := ini.ReadString('General', 'MSAA_cbHTML', 'HTML');
            mnuIA2.Caption := ini.ReadString('General', 'MSAA_cbIA2', 'IA2');
            mnuUIA.Caption := ini.ReadString('General', 'MSAA_cbUIA', 'UIAutomation');
            mnuMSAA.Hint := ini.ReadString('General', 'MSAA_cbMSAAHint', '');
            mnuARIA.Hint := ini.ReadString('General', 'MSAA_cbARIAHint', '');
            mnuHTML.Hint := ini.ReadString('General', 'MSAA_cbHTMLHint', '');
            mnuIA2.Hint := ini.ReadString('General', 'MSAA_cbIA2Hint', '');
            mnuUIA.Hint := ini.ReadString('General', 'MSAA_cbUIAHint', '');
            None := ini.ReadString('General', 'none', '(none)');
            ConvErr := ini.ReadString('General', 'MSAA_ConvertError', 'Format is invalid: %s');
            Msg := ini.ReadString('General', 'MSAA_HookIsFailed', 'SetWinEventHook is Failed!!');

            sTrue := ini.ReadString('General', 'true', 'True');
            sFalse := ini.ReadString('General', 'false', 'False');
            mnuTVSave.Caption := ini.ReadString('General', 'mnuSave', '&Save');
            mnuTVSAll.Caption := ini.ReadString('General', 'mnuSaveAll', '&All items');
            mnuTVSSel.Caption := ini.ReadString('General', 'mnuSaveSelect', '&Selected items');

            mnuTVOpen.Caption := ini.ReadString('General', 'mnuOpen', '&Open in Browser');
            mnuTVOAll.Caption := ini.ReadString('General', 'mnuOpenAll', '&All items');
            mnuTVOSel.Caption := ini.ReadString('General', 'mnuOpenSelect', '&Selected items');

            mnuSave.Caption := ini.ReadString('General', 'mnuSave', '&Save');
            mnuSAll.Caption := ini.ReadString('General', 'mnuSaveAll', '&All items');
            mnuSSel.Caption := ini.ReadString('General', 'mnuSaveSelect', '&Selected items');

            mnuOpenB.Caption := ini.ReadString('General', 'mnuOpen', '&Open in Browser');
            mnuOAll.Caption := ini.ReadString('General', 'mnuOpenAll', '&All items');
            mnuOSel.Caption := ini.ReadString('General', 'mnuOpenSelect', '&Selected items');

            mnuSelMode.Caption := ini.ReadString('General', 'mnuSelMode', 'S&elect Mode');

            tbFocus.Caption := ini.ReadString('General', 'MSAA_tbFocusName', 'Focus');
            tbCursor.Caption := ini.ReadString('General', 'MSAA_tbCursorName', 'Cursor');
            tbRectAngle.Caption := ini.ReadString('General', 'MSAA_tbRectAngleName', 'Highlight Rectangle');
            tbShowtip.Caption :=  ini.ReadString('General', 'MSAA_tbBalloonName', 'Balloon tip');
            tbCopy.Caption := ini.ReadString('General', 'MSAA_tbCopyName', 'Copy');
            tbOnlyFocus.Caption := ini.ReadString('General', 'MSAA_tbFocusOnlyName', 'Focus only');
            tbParent.Caption := ini.ReadString('General', 'MSAA_tbParentName', 'Parent');
            tbChild.Caption := ini.ReadString('General', 'MSAA_tbChildName', 'Child');
            tbPrevS.Caption := ini.ReadString('General', 'MSAA_tbPrevSName', 'Previous');
            tbNextS.Caption := ini.ReadString('General', 'MSAA_tbNextSName', 'Next');
            tbHelp.Caption := ini.ReadString('General', 'MSAA_tbHelpName', 'Help');
            tbMSAAMode.Caption := ini.ReadString('General', 'MSAA_tbMSAAModeName', 'MSAA');
            tbRegister.Caption := ini.ReadString('General', 'MSAA_tbMSAASetName', 'Settings');

            mnuReg.Caption := ini.ReadString('General', 'MSAA_mnuReg', '&Register');
            mnuUnReg.Caption := ini.ReadString('General', 'MSAA_mnuUnreg', '&Unregister');
            mnuReg.Hint := ini.ReadString('General', 'MSAA_mnuRegHint', 'Register IAccessible2Proxy.dll');
            mnuUnReg.Hint := ini.ReadString('General', 'MSAA_mnuUnregHint', 'Unregister IAccessible2Proxy.dll');

            Memo1.AccName := ini.ReadString('General', 'MSAA_CodeEdit', 'Code');
            Memo1.Hint := ini.ReadString('General', 'MSAA_CodeEditHint', '');
            Memo1.AccDesc := GetLongHint(Memo1.Hint);
            TreeList1.Header.Columns[0].Text := ini.ReadString('General', 'MSAA_tlColumn1', 'Intarfece');
            TreeList1.Header.Columns[1].Text := ini.ReadString('General', 'MSAA_tlColumn2', 'Value');
            TreeList1.AccessibleName := PChar(ini.ReadString('General', 'MSAA_ListName', 'Result List'));
            d := PChar(ini.ReadString('General', 'MSAA_ListName', 'Result List'));
            SetWindowText(TreeList1.Handle, PWideChar(d));
            TreeList1.Hint := ini.ReadString('General', 'MSAA_ListHint', '');
            //TreeList1.AccDesc := GetLongHint(TreeList1.Hint);
            TreeView1.AccName := PChar(ini.ReadString('General', 'MSAA_TVName', 'Accessibility Tree'));
            TreeView1.Hint := ini.ReadString('General', 'MSAA_TVHint', '');
            TreeView1.AccDesc := GetLongHint(TreeView1.Hint);
            UIA_fail := ini.ReadString('General', 'UIA_CreationError', 'CoCreateInstance failed. ');

            PB1.Hint :=  ini.ReadString('General', 'Collapse_TV_Hint', 'Collapse(or Expand) Button for TreeView');
            PB2.Hint :=  ini.ReadString('General', 'Collapse_TL_Hint', 'Collapse(or Expand) Button for TreeList');
            PB3.Hint :=  ini.ReadString('General', 'Collapse_CE_Hint', 'Collapse(or Expand) Button for CodeEdit');
            PB1.AccDesc := PB1.Hint;
            PB2.AccDesc := PB2.Hint;
            PB3.AccDesc := PB3.Hint;
            acTVCol.Caption := ini.ReadString('General', 'mnuTreeView', '&TreeView');
            acTLCol.Caption := ini.ReadString('General', 'mnuTreeList', 'Tree&List');
            acMMCol.Caption := ini.ReadString('General', 'mnuCodeEdit', '&CodeEdit');
            mnuColl.Caption := ini.ReadString('General', 'mnuCollapse', '&Collapse');

            mnuTVcont.Caption := ini.ReadString('General', 'mnuTVContent', '&Treeview contents');
            mnuTarget.Caption := ini.ReadString('General', 'mnuTarget', '&Target only');
            mnuAll.Caption := ini.ReadString('General', 'mnuAll', '&All related objects');

            rType := ini.ReadString('IA2', 'RelationType', 'Relation Type');
            rTarg := ini.ReadString('IA2', 'RelationTargets', 'Relation Targets');
            mnuBln.Caption := ini.ReadString('General', 'mnubln', 'Balloon tip(&B)');
            mnublnMSAA.Caption :=  ini.ReadString('General', 'mnublnMSAA', '&MS Active Accessibility');
            mnublnIA2.Caption :=  ini.ReadString('General', 'mnublnIA2', '&IAccessible2');
            mnublnCode.Caption :=  ini.ReadString('General', 'mnublnCode', '&Source code');

            ShowSrcLen := ini.ReadInteger('General', 'ShowSrcLen', 2000);
            ClsNames.CommaText := ini.ReadString('General', 'ClassNames', '"mozillawindowclass","chrome_renderwidgethosthwnd","mozillawindowclass","chrome_widgetwin_0","chrome_widgetwin_1"');



            sFilter :=  ini.ReadString('General', 'SaveDLGFilter', 'HTML File|*.htm*|Text File|*.txt|All|*.*');

            sHTML := ini.ReadString('HTML', 'HTML', 'HTML');
            HTMLs[0, 0] := ini.ReadString('HTML', 'Element_name', 'Element Name');
            HTMLs[1, 0] := ini.ReadString('HTML', 'Attributes', 'Attributes');
            HTMLs[2, 0] := ini.ReadString('HTML', 'Code', 'Code');
            sTypeIE := '(' + ini.ReadString('HTML', 'outerHTML', 'outerHTML') + ')';
            HTMLs[2, 0] := HTMLs[2, 0] + StypeIE;

            HTMLsFF[0, 0] := HTMLs[0, 0];
            HTMLsFF[1, 0] := HTMLs[1, 0];
            HTMLsFF[2, 0] := HTMLs[2, 0];
            sTxt := ini.ReadString('HTML', 'Text', 'Text');
            sTypeFF := '(' + ini.ReadString('HTML', 'innerHTML', 'innerHTML') + ')';

            sAria := ini.ReadString('ARIA', 'ARIA', 'ARIA');
            ARIAs[0, 0] := ini.ReadString('ARIA', 'Role', 'Role');
            ARIAs[1, 0] := ini.ReadString('ARIA', 'Attributes', 'Attributes');

            Roles[0] := ini.ReadString('IA2', 'IA2_ROLE_CANVAS', 'canvas');
            Roles[1] := ini.ReadString('IA2', 'IA2_ROLE_CAPTION', 'Caption');
            Roles[2] := ini.ReadString('IA2', 'IA2_ROLE_CHECK_MENU_ITEM', 'Check menu item');
            Roles[3] := ini.ReadString('IA2', 'IA2_ROLE_COLOR_CHOOSER', 'Color chooser');
            Roles[4] := ini.ReadString('IA2', 'IA2_ROLE_DATE_EDITOR', 'Date editor');
            Roles[5] := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_ICON', 'Desktop icon');
            Roles[6] := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_PANE', 'Desktop pane');
            Roles[7] := ini.ReadString('IA2', 'IA2_ROLE_DIRECTORY_PANE', 'Directory pane');
            Roles[8] := ini.ReadString('IA2', 'IA2_ROLE_EDITBAR', 'Editbar');
            Roles[9] := ini.ReadString('IA2', 'IA2_ROLE_EMBEDDED_OBJECT', 'Embedded object');
            Roles[10] := ini.ReadString('IA2', 'IA2_ROLE_ENDNOTE', 'Endnote');
            Roles[11] := ini.ReadString('IA2', 'IA2_ROLE_FILE_CHOOSER', 'File chooser');
            Roles[12] := ini.ReadString('IA2', 'IA2_ROLE_FONT_CHOOSER', 'Font chooser');
            Roles[13] := ini.ReadString('IA2', 'IA2_ROLE_FOOTER', 'Footer');
            Roles[14] := ini.ReadString('IA2', 'IA2_ROLE_FOOTNOTE', 'Footnote');
            Roles[15] := ini.ReadString('IA2', 'IA2_ROLE_FORM', 'Form');
            Roles[16] := ini.ReadString('IA2', 'IA2_ROLE_FRAME', 'Frame');
            Roles[17] := ini.ReadString('IA2', 'IA2_ROLE_GLASS_PANE', 'Glass pane');
            Roles[18] := ini.ReadString('IA2', 'IA2_ROLE_HEADER', 'Header');
            Roles[19] := ini.ReadString('IA2', 'IA2_ROLE_HEADING', 'Heading');
            Roles[20] := ini.ReadString('IA2', 'IA2_ROLE_ICON', 'Icon');
            Roles[21] := ini.ReadString('IA2', 'IA2_ROLE_IMAGE_MAP', 'Image map');
            Roles[22] := ini.ReadString('IA2', 'IA2_ROLE_INPUT_METHOD_WINDOW', 'Input method window');
            Roles[23] := ini.ReadString('IA2', 'IA2_ROLE_INTERNAL_FRAME', 'Internal frame');
            Roles[24] := ini.ReadString('IA2', 'IA2_ROLE_LABEL', 'Label');
            Roles[25] := ini.ReadString('IA2', 'IA2_ROLE_LAYERED_PANE', 'Layered pane');
            Roles[26] := ini.ReadString('IA2', 'IA2_ROLE_NOTE', 'Note');
            Roles[27] := ini.ReadString('IA2', 'IA2_ROLE_OPTION_PANE', 'Option pane');
            Roles[28] := ini.ReadString('IA2', 'IA2_ROLE_PAGE', 'Role pane');
            Roles[29] := ini.ReadString('IA2', 'IA2_ROLE_PARAGRAPH', 'Paragraph');
            Roles[30] := ini.ReadString('IA2', 'IA2_ROLE_RADIO_MENU_ITEM', 'Radio menu item');
            Roles[31] := ini.ReadString('IA2', 'IA2_ROLE_REDUNDANT_OBJECT', 'Redundant object');
            Roles[32] := ini.ReadString('IA2', 'IA2_ROLE_ROOT_PANE', 'Root pane');
            Roles[33] := ini.ReadString('IA2', 'IA2_ROLE_RULER', 'Ruler');
            Roles[34] := ini.ReadString('IA2', 'IA2_ROLE_SCROLL_PANE', 'Scroll pane');
            Roles[35] := ini.ReadString('IA2', 'IA2_ROLE_SECTION', 'Section');
            Roles[36] := ini.ReadString('IA2', 'IA2_ROLE_SHAPE', 'Shape');
            Roles[37] := ini.ReadString('IA2', 'IA2_ROLE_SPLIT_PANE', 'Split pane');
            Roles[38] := ini.ReadString('IA2', 'IA2_ROLE_TEAR_OFF_MENU', 'Tear off menu');
            Roles[39] := ini.ReadString('IA2', 'IA2_ROLE_TERMINAL', 'Terminal');
            Roles[40] := ini.ReadString('IA2', 'IA2_ROLE_TEXT_FRAME', 'Text frame');
            Roles[41] := ini.ReadString('IA2', 'IA2_ROLE_TOGGLE_BUTTON', 'Toggle button');
            Roles[42] := ini.ReadString('IA2', 'IA2_ROLE_VIEW_PORT', 'View port');

            IA2Sts[0] := ini.ReadString('IA2', 'IA2_STATE_ACTIVE', 'Active');
            IA2Sts[1] := ini.ReadString('IA2', 'IA2_STATE_ARMED', 'Armed');
            IA2Sts[2] := ini.ReadString('IA2', 'IA2_STATE_DEFUNCT', 'Defunct');
            IA2Sts[3] := ini.ReadString('IA2', 'IA2_STATE_EDITABLE', 'Editable');
            IA2Sts[4] := ini.ReadString('IA2', 'IA2_STATE_HORIZONTAL', 'Horizontal');
            IA2Sts[5] := ini.ReadString('IA2', 'IA2_STATE_ICONIFIED', 'Iconified');
            IA2Sts[6] := ini.ReadString('IA2', 'IA2_STATE_INVALID_ENTRY', 'Invalid entry');
            IA2Sts[7] := ini.ReadString('IA2', 'IA2_STATE_MANAGES_DESCENDANTS', 'Manages descendants');
            IA2Sts[8] := ini.ReadString('IA2', 'IA2_STATE_MODAL', 'Modal');
            IA2Sts[9] := ini.ReadString('IA2', 'IA2_STATE_MULTI_LINE', 'Multi line');
            IA2Sts[10] := ini.ReadString('IA2', 'IA2_STATE_OPAQUE', 'Opaque');
            IA2Sts[11] := ini.ReadString('IA2', 'IA2_STATE_REQUIRED', 'Required');
            IA2Sts[12] := ini.ReadString('IA2', 'IA2_STATE_SELECTABLE_TEXT', 'Selectable text');
            IA2Sts[13] := ini.ReadString('IA2', 'IA2_STATE_SINGLE_LINE', 'Single line');
            IA2Sts[14] := ini.ReadString('IA2', 'IA2_STATE_STALE', 'Stale');
            IA2Sts[15] := ini.ReadString('IA2', 'IA2_STATE_SUPPORTS_AUTOCOMPLETION', 'Supports autocompletion');
            IA2Sts[16] := ini.ReadString('IA2', 'IA2_STATE_TRANSIENT', 'Transient');
            IA2Sts[17] := ini.ReadString('IA2', 'IA2_STATE_VERTICAL', 'Vertical');

            {acFocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbFocus', 'F4'));
            acCursor.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbCursor', 'F5'));
            acRect.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbRectAngle', 'F6'));
            acShowtip.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbBalloon', 'F3'));
            acCopy.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbCopy', 'F7'));
            acOnlyfocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbFocusOnly', 'F8'));
            acParent.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbParent', 'F9'));
            acChild.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbChild', 'F10'));
            acPrevS.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbPrevS', 'F11'));
            acNextS.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbNextS', 'F12'));
            acHelp.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbHelp', 'F1'));
            acMSAAMode.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbMSAAMode', ''));
            acSetting.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'tbSetting', ''));
            acTreeFocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'sfTreeView', 'Shift+F1'));
            acListFocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'sfListView', 'Shift+F2'));
            acMemoFocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'sfTextBox', 'Shift+F3'));
            acTriFocus.ShortCut := TextToShortcut(ini.ReadString('Shortcut', 'sf3ctrls', 'Shift+F4')); }

            QSFailed := ini.ReadString('IA2', 'QS_Failed', 'Query Service Failed');
            Err_Inter := ini.ReadString('General', 'Error_Interface', 'Interface not available');
        finally
            Result := UIA_fail;
            ini.Free;
        end;
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
    sList: TStringList;
begin
    hHook := SetWinEventHook(EVENT_MIN,EVENT_MAX, 0, @WinEventProc, 0, 0,
        WINEVENT_OUTOFCONTEXT or WINEVENT_SKIPOWNPROCESS );
    Msg := 'SetWinEventHook is Failed!!';
    //flgMSAA, flgIA2, flgUIA: integer;
    flgMSAA := 125;
    flgIA2 := 119;
    flgUIA := 50332680;
    flgUIA2 := 0;
    iDefIndex := -1;
    if fileexists(sPath) then
    begin
      sList := TStringList.Create;
      try
        sList.LoadFromFile(sPath);
        if sList.Encoding <> TEncoding.Unicode then
          sList.SaveToFile(sPath, TEncoding.Unicode);
      finally
        sList.Free;
      end;
    end;

    //iAccS := nil;
    ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
    try
        if mnuLang.Visible then
        begin
            d := Ini.ReadString('Settings', 'LangFile', 'Default.ini');
            TransPath := TransDir + d;
        end;

        LoadLang;
        Width := ini.ReadInteger('Settings', 'Width', 335);
        Height := ini.ReadInteger('Settings', 'Height', 500);
        DefW := Width;
        DefH := Height;
        Width := DoubleToInt(DefW * ScaleX);
        Height := DoubleToInt(DefH * ScaleY);
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
        DefFont := Font.Size;
        Font.Size := DoubleToInt(DefFont * ScaleY);

        mnuMSAA.Checked := ini.ReadBool('Settings', 'vMSAA', True);
        mnuARIA.Checked := ini.ReadBool('Settings', 'vARIA', True);
        mnuHTML.Checked := ini.ReadBool('Settings', 'vHTML', True);
        mnuIA2.Checked := ini.ReadBool('Settings', 'vIA2', True);
        mnuUIA.Checked := ini.ReadBool('Settings', 'vUIA', True);

        mnublnMSAA.Checked := ini.ReadBool('Settings', 'bMSAA', True);
        mnublnIA2.Checked := ini.ReadBool('Settings', 'bIA2', True);
        mnublnCode.Checked := ini.ReadBool('Settings', 'bCode', True);

        acFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acFocus.Name, 115));
        acCursor.ShortCut := Word(ini.ReadInteger('Shortcut', acCursor.Name, 116));
        acRect.ShortCut := Word(ini.ReadInteger('Shortcut', acRect.Name, 117));
        acShowtip.ShortCut := Word(ini.ReadInteger('Shortcut', acShowtip.Name, 114));
        acCopy.ShortCut := Word(ini.ReadInteger('Shortcut', acCopy.Name, 118));
        acOnlyfocus.ShortCut := Word(ini.ReadInteger('Shortcut', acOnlyfocus.Name, 119));
        acParent.ShortCut := Word(ini.ReadInteger('Shortcut', acParent.Name, 120));
        acChild.ShortCut := Word(ini.ReadInteger('Shortcut', acChild.Name, 121));
        acPrevS.ShortCut := Word(ini.ReadInteger('Shortcut', acPrevS.Name, 122));
        acNextS.ShortCut := Word(ini.ReadInteger('Shortcut', acNextS.Name, 123));
        acHelp.ShortCut := Word(ini.ReadInteger('Shortcut', acHelp.Name, 112));
        acMSAAMode.ShortCut := Word(ini.ReadInteger('Shortcut', acMSAAMode.Name, 8308));
        acSetting.ShortCut := Word(ini.ReadInteger('Shortcut', acSetting.Name, 8309));
        acTreeFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acTreeFocus.Name, 8304));
        acListFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acListFocus.Name, 8305));
        acMemoFocus.ShortCut := Word(ini.ReadInteger('Shortcut', acMemoFocus.Name, 8306));
        ac3ctrls.ShortCut := Word(ini.ReadInteger('Shortcut', ac3ctrls.Name, 8307));

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
            ini := TMemIniFile.Create(TransDir + LangList[i], TEncoding.UTF8);
            try
                mItem := MainMenu1.CreateMenuItem;
                mItem.Caption := ini.ReadString('General', 'Language', 'English');
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
                ini.Free;
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

    end
    else
    begin
      UIAuto.AddFocusChangedEventHandler(nil, self);
      hr := UIAuto.ElementFromHandle(pointer(handle), uiEle);
      if (hr = 0) and (Assigned(uiEle)) then
      	uiEle.Get_CurrentProcessId(iPID);
			uiEle := nil;
    end;
    TreeList1.Header.Columns[0].Width := TreeList1.ClientWidth div 2;
    TreeList1.Header.Columns[1].Width := TreeList1.ClientWidth div 2;
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
    SizeChange;
    Toolbar1.Focused;
end;



procedure TwndMSAAV.mnuMSAAClick(Sender: TObject);
begin
	if (not mnuMSAA.Checked) and (not mnuARIA.Checked) and (not mnuHTML.Checked)
		and (not mnuIA2.Checked) and (not mnuUIA.Checked) then
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

procedure TwndMSAAV.RecursiveTV(cNode: TTreeNode; var HTML: string; var iCnt: integer; ForSel: boolean = false);
var
	i: integer;
  tab, temp, ia2t: string;
  Role: string;
  ws: widestring;
  ovChild: OleVariant;
  iRole :integer;
  ovValue: OleVariant;

  function GetLIContents(Acc: IAccessible; Child: integer; pNode: TTreeNode ): string;
  var
  	Res: string;
    PC:PChar;
  begin
  	Role := '';
		ovChild := Child;
		Acc.Get_accRole(ovChild, ovValue);
		iRole := TVarData(ovValue).VInteger;
		try
			if VarHaveValue(ovValue) and VarIsNumeric(ovValue) then
			begin

				if pNode.Text = '' then
				begin
					Acc.Get_accName(ovChild, ws);
					if ws = '' then
						ws := None;
					ws := ReplaceStr(ws, '<', '&lt;');
					ws := ReplaceStr(ws, '>', '&gt;');
					PC := StrAlloc(255);
					GetRoleTextW(ovValue, PC, StrBufSize(PC));
					Role := PC;
					StrDispose(PC);
					pNode.Text := ws + ' - ' + Role;
				end;
			end;
			if (iRole = ROLE_SYSTEM_STATICTEXT) or (iRole = ROLE_SYSTEM_TEXT) or
				(iRole = ROLE_SYSTEM_APPLICATION) or (iRole = ROLE_SYSTEM_WINDOW) then
			begin
				ws := ReplaceStr(pNode.Text, '<', '&lt;');
				ws := ReplaceStr(ws, '>', '&gt;');
				Res := Res + #13#10#9 + tab + '<li>' + ws;
			end
			else
			begin

				try
					Res := Res + #13#10 + tab + '<li class="element">' + #13#10 + tab +
						'<input type="checkbox" id="disclosure' + inttostr(iCnt) +
						'" title="check to display details below" aria-controls="x-details'
						+ inttostr(iCnt) + '"> ';
					Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt)
						+ '"><strong>' + UpperCase(GetEleName(TTreeData(pNode.Data^).Acc)) +
						'</strong></label> ';
					Res := Res + #13#10 + tab + '<section id="x-details' +
						inttostr(iCnt) + '">';
					Res := Res + #13#10 + tab + '<ul>' + #13#10 + tab + #9 +
						'<li class="API">';
					Res := Res + MSAAText4HTML(Acc, tab + #9) + #13#10 + tab + #9
						+ '</li>';

					ia2t := IA2Text4HTML(Acc, tab + #9);
					if ia2t <> '' then
					begin
						Res := Res + #13#10 + tab + #9 + '<li class="API">';
						Res := Res + ia2t + #13#10 + tab + #9 + '</li>';
					end;
				finally
					Res := Res + #13#10 + #9 + '</ul>' + #13#10 + tab + '</section>';
					inc(iCnt);
				end;

			end;

		finally
			Result := Res;
		end;
  end;

  function GetLIC_UIA(pNode: TTreeNode): string;
  var
  	Res: string;
    tempEle: IUIAUTOMATIONELEMENT;
  begin
    tempEle := uiEle;
  	try
    	uiEle := TTreeData(pNode.Data^).uiEle;
			Res := Res + #13#10 + tab + '<li class="element">' + #13#10 + tab +
				'<input type="checkbox" id="disclosure' + inttostr(iCnt) +
				'" title="check to display details below" aria-controls="x-details' +
				inttostr(iCnt) + '"> ';
			Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt) +
				'"><strong>' + UpperCase(pNode.Text) +
				'</strong></label> ';
			Res := Res + #13#10 + tab + '<section id="x-details' +
				inttostr(iCnt) + '">';
			Res := Res + #13#10 + tab + '<ul>' + #13#10 + tab + #9 +
				'<li class="API">';
			Res := Res + UIAText(True, tab + #9) + #13#10 + tab + #9 + '</li>';
      Res := Res + #13#10 + #9 + '</ul>' + #13#10 + tab + '</section>';
		finally
      Result := Res;
			inc(iCnt);
      uiEle := tempEle;
		end;
  end;

begin
	if not Assigned(cNode.Data) then
  	Exit;

	tab := #9;
	if cNode.Level > 0 then
  begin
		for i := 1 to cNode.Level do
  		tab := tab + #9;
  end;
  if ForSel then
  begin
  	if cNode.Selected then
    begin
			if not mnutvUIA.Checked then
      	HTML := HTML + GetLIContents(TTreeData(cNode.Data^).Acc, TTreeData(cNode.Data^).iID, cNode)
      else
       	HTML := HTML + GetLIC_UIA(cNode);
    end;
  end
  else
	begin

		if Pagecontrol1.ActivePageIndex = 0 then
			HTML := HTML + GetLIContents(TTreeData(cNode.Data^).Acc,
				TTreeData(cNode.Data^).iID, cNode)
		else
			HTML := HTML + GetLIC_UIA(cNode);
	end;
  if cNode.HasChildren then
  begin
  	temp := '';
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
				RecursiveTV(cNode.Item[i], temp, iCnt, Forsel);
      end
      else
      begin
      	if ForSel then
  			begin
  				if cNode.Item[i].Selected then
          begin
          	if not mnutvUIA.Checked then
       				temp := temp + #13#10 + tab + #9 + GetLIContents(TTreeData(cNode.Item[i].Data^).Acc, TTreeData(cNode.Item[i].Data^).iID, cNode.Item[i]) + '</li>'
            else
            	temp := temp + #13#10 + tab + #9 + GetLIC_UIA(cNode.Item[i]) + '</li>';
          end;
        end
  			else
				begin
					if Pagecontrol1.ActivePageIndex = 0 then
						temp := temp + #13#10 + tab + #9 +
							GetLIContents(TTreeData(cNode.item[i].Data^).Acc,
							TTreeData(cNode.item[i].Data^).iID, cNode.item[i]) + '</li>'
					else
						temp := temp + #13#10 + tab + #9 + GetLIC_UIA(cNode.Item[i]) + '</li>';
				end;
      end;
    end;
    if temp <> '' then
    begin
      HTML := HTML + #13#10 + tab + '<ul>' + temp + #13#10 + tab + '</ul>';
    end;
  end;


end;

procedure TwndMSAAV.mnuSAllClick(Sender: TObject);
begin
	mnuTVSAllClick(self);
end;

procedure TwndMSAAV.mnuSelModeClick(Sender: TObject);
begin
  bSelMode := not bSelMode;
  mnuSelmode.Checked := bSelMode;
end;

procedure TwndMSAAV.mnuSSelClick(Sender: TObject);
begin
	mnuTVSSelClick(self);
end;

function GetTemp(var S: String): Boolean;
var
  Len: Integer;
begin
  Len := Windows.GetTempPath(0, nil);
  if Len > 0 then
  begin
    SetLength(S, Len);
    Len := Windows.GetTempPath(Len, PChar(S));
    SetLength(S, Len);
    S := SysUtils.IncludeTrailingPathDelimiter(S);
    Result := Len > 0;
  end else
    Result := False;
end;

function TwndMSAAV.GetEleName(Acc: IAccessible): string;
var
  isp: iserviceprovider;
  hr: hResult;
  isEle: ISimpleDOMNode;
  PC, PC2:PChar;
  SI: Smallint;
  PU, PU2: PUINT;
  WD: Word;
  iEle: IHTMLElement;
begin
  Result := '';
  hr := Acc.QueryInterface(IID_IServiceProvider, iSP);
  if (hr = 0) and Assigned(iSP) then
	begin
		hr := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
		if (hr = 0) and Assigned(iEle) then
		begin
			Result := iEle.tagName;
		end
    else
    begin
    	hr := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle);
      if (hr = 0) and Assigned(isEle) then
      begin
        isEle.get_nodeInfo(PC, SI, PC2, PU, PU2, WD);
        Result := PC;
      end;
    end;
	end;
end;

function TwndMSAAV.GetTVAllItems: string;
var
	d, temp, sTitle: string;
  sList: TStringList;
  i: integer;
  isp: iserviceprovider;
  hr: hResult;
  isEle, pEle: ISimpleDOMNode;
  isDoc: ISimpleDOMDocument;
  aPC : pWidechar;
  iEle: IHTMLElement;
  iDoc2: IHTMLDocument2;
  tTreeV: TAccTreeView;
begin
	if PageControl1.ActivePageIndex = 1 then
  begin
		sTitle := 'aViewer UIA Tree';
    tTreeV := tbUIA;
  end
	else
	begin
  	tTreeV := TreeView1;
		hr := iAcc.QueryInterface(IID_IServiceProvider, iSP);
		if (hr = 0) and Assigned(iSP) then
		begin
			hr := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
			if (hr = 0) and Assigned(iEle) then
			begin
				hr := iEle.Document.QueryInterface(IID_IHTMLDOCUMENT2, iDoc2);
				if (hr = 0) and Assigned(iDoc2) then
				begin
					sTitle := iDoc2.Title;
				end;
			end;
		end
		else
		begin
			hr := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle);
			if (hr = 0) and Assigned(isEle) then
			begin

				hr := isEle.QueryInterface(IID_ISIMPLEDOMDOCUMENT, isDoc);
				i := 0;
				while not Assigned(isDoc) do
				begin
					isEle.get_parentNode(pEle);
					if not Assigned(pEle) then
						break;
					isEle := pEle;

					hr := isEle.QueryInterface(IID_ISIMPLEDOMDOCUMENT, isDoc);
					inc(i);
					if i > 1000 then
						break;
				end;
				if SUCCEEDED(hr) and Assigned(isDoc) then
				begin
					isDoc.get_title(aPC);
					sTitle := aPC;
				end;
			end;
		end;
	end;

  i := 0;
  RecursiveTV(tTreeV.Items.Item[0], d, i);
  d := '<ul>' + d  + #13#10 + '</li>' + #13#10 + '</ul>';
  if PageControl1.ActivePageIndex = 1 then
  	d := '<h1>UIAutomation Tree</h1>' + #13#10 + d
  else
  	d := '<h1>Accessibility Tree</h1>' + #13#10 + d;

  sList := TStringList.Create;
  try

    temp := APPDir + 'output.html';
    if FileExists(temp) then
    begin
      sList.LoadFromFile(temp, TEncoding.UTF8);
      sList.Text := StringReplace(sList.Text, '%title%', sTitle, [rfReplaceAll, rfIgnoreCase]);
      sList.Text := StringReplace(sList.Text, '%contents%', d, [rfReplaceAll, rfIgnoreCase]);
    end
    else
      sList.Text := d;
    Result := sList.Text;
  finally
    sList.Free;
  end;
end;

procedure TwndMSAAV.mnuOAllClick(Sender: TObject);
begin
  mnuTVOAllClick(self);
end;

function TwndMSAAV.SaveHTMLDLG(initDir: string; var FName: string): boolean;
var
  ofn: TOpenFileName;
  szFile: array[0..MAX_PATH] of Char;
begin
  Result := False;
  FillChar(ofn, SizeOf(TOpenFileName), 0);

  with ofn do
  begin
    lStructSize := SizeOf(TOpenFileName);
    hwndOwner := Handle;
    Flags := OFN_OVERWRITEPROMPT or OFN_HIDEREADONLY;
    lpstrFile := szFile;
    nMaxFile := SizeOf(szFile);
    if (initDir <> '') then
      lpstrInitialDir := PChar(initDir);
    StrPCopy(lpstrFile, FName);
    lpstrFilter := PChar(ReplaceStr(sFilter, '|', #0)+#0#0);;
    nFilterIndex := 1;
    lpstrDefExt := 'html';
  end;

  if GetSaveFileName(ofn) then
  begin
  	Result := True;
  	FName := StrPas(szFile);
  end;
end;

procedure TwndMSAAV.mnuTVOAllClick(Sender: TObject);
var
  temp: string;
  sList: TStringList;
begin
  if TreeView1.Items.Count > 0 then
  begin
    if GetTemp(temp) then
    begin
      sList := TStringList.Create;
      try
        sList.Text := GetTVAllItems;
        sList.SaveToFile(temp+'aviewer.html', TEncoding.UTF8);
        ShellExecute(Handle, 'open', PWideChar(temp+'aviewer.html'), nil, nil, SW_SHOW);
      finally
        sList.Free;
      end;
    end;
  end;
end;


procedure TwndMSAAV.mnuTVSAllClick(Sender: TObject);
var
  sList: TStringList;
  FName: string;
begin
    if TreeView1.Items.Count > 0 then
    begin
      sList := TStringList.Create;
      bPFunc := True;
      try

        sList.Text := GetTVAllItems;
        if SaveHTMLDLG('', FName) then
        begin
        	sList.SaveToFile(FName, TEncoding.UTF8);
        end;

      finally
        sList.Free;
        bPFunc := False;
      end;
    end;

end;

function TwndMSAAV.GetTVSelItems: string;
var
	d, temp, sTitle: string;
  sList: TStringList;
  //i: integer;
  	i,iCnt: integer;
  tab{, temp}, ia2t: string;
  Role: string;
  ws: widestring;
  ovChild: OleVariant;
  iRole :integer;
  ovValue: OleVariant;
  hr: hResult;
  iEle: IHTMLElement;
  iDoc2: IHTMLDocument2;
  isp: iserviceprovider;
  isEle, pEle: ISimpleDOMNode;
  isDoc: ISimpleDOMDocument;
  aPC : pWidechar;
  tTreeV: TAccTreeView;
  function GetLIContents(Acc: IAccessible; Child: integer; pNode: TTreeNode ): string;
  var
  	Res: string;
    PC:PChar;
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
  		Res := Res + #13#10#9 + tab + '<li>' + ws;
  	end
  	else
  	begin

    	try
  		Res := Res + #13#10 + tab + '<li class="element">' + #13#10 + tab + '<input type="checkbox" id="disclosure' + inttostr(iCnt) + '" title="check to display details below" aria-controls="x-details' + inttostr(iCnt) + '"> ';
    	Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt) + '"><strong>' + UpperCase(GetElename(TTreeData(pNode.Data^).Acc)) + '</strong></label> ';
    	Res := Res + #13#10 + tab + '<section id="x-details' + inttostr(iCnt) + '">';
      Res := Res + #13#10#9 + tab + '<ul>' + #13#10 + tab + #9 + '<li class="API">';
      Res := Res + MSAAText4HTML(Acc, tab + #9) + #13#10 + tab + #9 + '</li>';
      ia2t := IA2Text4HTML(Acc, tab + #9);
      if ia2t <> '' then
      begin
      	Res := Res + #13#10 + tab + #9 + '<li class="API">';
      	Res := Res + ia2t + #13#10 + tab + #9 + '</li>';
      end;

      finally
      	Res := Res + #13#10 + #9 + '</ul>' + #13#10 + tab + '</section>';
      	Inc(iCnt);
      end;

  	end;

    finally
    	Result := Res;
    end;
  end;

  function GetLIC_UIA(pNode: TTreeNode): string;
  var
  	Res: string;
    tempEle: IUIAUTOMATIONELEMENT;
  begin
    tempEle := uiEle;
  	try
    	uiEle := TTreeData(pNode.Data^).uiEle;
			Res := Res + #13#10 + tab + '<li class="element">' + #13#10 + tab +
				'<input type="checkbox" id="disclosure' + inttostr(iCnt) +
				'" title="check to display details below" aria-controls="x-details' +
				inttostr(iCnt) + '"> ';
			Res := Res + #13#10 + tab + '<label for="disclosure' + inttostr(iCnt) +
				'"><strong>' + UpperCase(pNode.Text) +
				'</strong></label> ';
			Res := Res + #13#10 + tab + '<section id="x-details' +
				inttostr(iCnt) + '">';
			Res := Res + #13#10 + tab + '<ul>' + #13#10 + tab + #9 +
				'<li class="API">';
			Res := Res + UIAText(True, tab + #9) + #13#10 + tab + #9 + '</li>';
      Res := Res + #13#10 + #9 + '</ul>' + #13#10 + tab + '</section>';
		finally
      Result := Res;
			inc(iCnt);
      uiEle := tempEle;
		end;
  end;
begin
  Result := '';
  if Pagecontrol1.ActivePageIndex = 1 then
  begin
		sTitle := 'aViewer UIA Tree' ;
    tTreeV := tbUIA;
  end
	else
	begin
  	tTreeV := TreeView1;
		hr := iAcc.QueryInterface(IID_IServiceProvider, iSP);
		if (hr = 0) and Assigned(iSP) then
		begin
			hr := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
			if (hr = 0) and Assigned(iEle) then
			begin
				hr := iEle.Document.QueryInterface(IID_IHTMLDOCUMENT2, iDoc2);
				if (hr = 0) and Assigned(iDoc2) then
				begin
					sTitle := iDoc2.Title;
				end;
			end;
		end
		else
		begin
			hr := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle);
			if (hr = 0) and Assigned(isEle) then
			begin

				hr := isEle.QueryInterface(IID_ISIMPLEDOMDOCUMENT, isDoc);
				i := 0;
				while not Assigned(isDoc) do
				begin
					isEle.get_parentNode(pEle);
					if not Assigned(pEle) then
						break;
					isEle := pEle;

					hr := isEle.QueryInterface(IID_ISIMPLEDOMDOCUMENT, isDoc);
					inc(i);
					if i > 1000 then
						break;
				end;
				if SUCCEEDED(hr) and Assigned(isDoc) then
				begin
					isDoc.get_title(aPC);
					sTitle := aPC;
				end;
			end;
		end;
	end;
  if (tTreeV.SelectionCount > 0) and (tTreeV.Items.Count > 0) then
    begin

      iCnt := 0;
      for i := 0 to tTreeV.SelectionCount - 1 do
      begin
      	if Pagecontrol1.ActivePageIndex = 0 then
					d := d + GetLIContents(TTreeData(tTreeV.Selections[i].Data^).Acc, TTreeData(tTreeV.Selections[i].Data^).iID, tTreeV.Selections[i]) + '</li>' + #13#10
      	else
          d := d + GetLIC_UIA(tTreeV.Selections[i]) + '</li>' + #13#10;
      end;
      d := '<ul>' + d + '</ul>';
      sList := TStringList.Create;
      try
      	temp := APPDir + 'output.html';
        if FileExists(temp) then
        begin
        	sList.LoadFromFile(temp, TEncoding.UTF8);
          sList.Text := StringReplace(sList.Text, '%title%', sTitle, [rfReplaceAll, rfIgnoreCase]);
          sList.Text := StringReplace(sList.Text, '%contents%', d, [rfReplaceAll, rfIgnoreCase]);
        end
        else
        	sList.Text := d;
        Result := sList.Text;
      finally
        sList.Free;
      end;
    end;
end;

procedure TwndMSAAV.mnuOSelClick(Sender: TObject);
begin
  mnuTVOSelClick(self);
end;

procedure TwndMSAAV.mnuTVOSelClick(Sender: TObject);
var
  temp: string;
  sList: TStringList;
begin
  if (TreeView1.SelectionCount > 0) and (TreeView1.Items.Count > 0) then
  begin
    if GetTemp(temp) then
    begin
      sList := TStringList.Create;
      try

        sList.Text := GetTVSelItems;
        sList.SaveToFile(temp+'aviewer.html', TEncoding.UTF8);
        ShellExecute(Handle, 'open', PWideChar(temp+'aviewer.html'), nil, nil, SW_SHOW);
      finally
        sList.Free;
      end;
    end;
  end;

end;

procedure TwndMSAAV.mnuTVSSelClick(Sender: TObject);
var
  sList: TStringList;
  FName: string;
begin
		//Save Treeview contents
    if (TreeView1.SelectionCount > 0) and (TreeView1.Items.Count > 0) then
    begin
      sList := TStringList.Create;
      bPFunc := True;
      try

      	sList.Text := GetTVSelItems;
        if SaveHTMLDLG('', FName) then
        begin
        	sList.SaveToFile(FName, TEncoding.UTF8);
        end;

      finally
        sList.Free;
        bPFunc := False;
      end;
    end;

end;


procedure TwndMSAAV.SizeChange;
var
	SZ: TSize;
	i, iHeight: integer;
  dBMP, mBMP, oriBMP: TBitmap;
  tpColor: TColor;
begin
		Font.Size := DoubleToInt(DefFont * ScaleY);
  	Width := DoubleToInt(DefW * ScaleX);
  	Height := DoubleToInt(DefH * ScaleY);
		GetTextExtentPoint(Canvas.Handle, PWideChar(Caption), Length(PwideChar(Caption)), SZ);
    TreeList1.Font := Font;
    TreeList1.Header.Font := Font;
    TreeList1.Header.Height := SZ.Height + 3;

    iHeight := DoubleToInt(16 * ScaleX);
    Imagelist4.Clear;
    ImageList4.Height := iHeight;
    ImageList4.Width := iHeight;
    for i := 0 to ImageList1.Count - 1 do
    begin
    	dBMP := TBitmap.Create;
      mBMP := TBitmap.Create;
      oriBMP := TBitmap.Create;
      try


        Imagelist1.GetBitmap(i, oriBMP);
        dBMP.PixelFormat :=pf24bit;
				mBMP.PixelFormat :=pf24bit;
  			dBMP.Width := iHeight;
  			dBMP.Height := iHeight;
  			mBMP.Width := iHeight;
  			mBMP.Height := iHeight;

  			tpColor := OriBMP.Canvas.Pixels[0, 0];

    		StretchBlt(dBMP.Canvas.Handle, 0, 0, dBMP.Width, dBMP.Height, OriBMP.Canvas.Handle, 0, 0, OriBMP.Width, OriBMP.Height, SRCCOPY);
  			StretchBlt(mBMP.Canvas.Handle, 0, 0, mBMP.Width, mBMP.Height, OriBMP.Canvas.Handle, 0, 0, OriBMP.Width, OriBMP.Height, SRCCOPY);
  			mBMP.Mask(tpColor);


  			ImageList4.Add(dBMP, mBMP);
      finally
      	dBMP.Free;
  			mBMP.Free;
        oriBMP.Free;
      end;
    end;
    toolbar1.Images := ImageList4;
    toolbar1.ButtonHeight := iHeight + 5;
    toolbar1.ButtonWidth := iHeight + 5;
end;

procedure TwndMSAAV.Splitter1Moved(Sender: TObject);
begin
    P2W := Panel2.Width;
end;

procedure TwndMSAAV.Splitter2Moved(Sender: TObject);
begin
    P4H := Panel4.Height;
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


procedure TwndMSAAV.PageControl1Change(Sender: TObject);
begin
	if PageControl1.ActivePageIndex = 0 then
  	acShowTip.Enabled := True
  else
  begin
  	acShowTip.Enabled := False;
  	if Assigned(WndTip) then
		begin
			WndTip.Visible := false;
		end;
  end;
  GetNaviState;
end;

procedure TwndMSAAV.PopupMenu2Popup(Sender: TObject);
begin

  if (TreeView1.Items.Count = 0) then
  begin
  	mnuSave.Enabled := false;
    mnuOpenB.Enabled := false;
    mnuSelMode.Enabled := False;
  end
  else
  begin
    mnuSave.Enabled := True;
    mnuOpenB.Enabled := True;
    mnuSelMode.Enabled := True;
    if TreeView1.SelectionCount = 0 then
    begin
  		mnuSSel.Enabled := false;
      mnuOSel.Enabled := false;
    end
  	else
    begin
    	mnuSSel.Enabled := True;
      mnuOSel.Enabled := True;
    end;
  end;
end;

procedure TwndMSAAV.ThHTMLDone(Sender: TObject);
begin
	sHTMLTxt := sHTML + #13#10 + HTMLs[0, 0] + ':' + #9 + HTMLs[0, 1] + #13#10 +
		HTMLs[1, 0] + ':' + #9 + HTMLs[1, 1] + #13#10 + HTMLs[2, 0] + ':' + #9 +
		HTMLs[2, 1] + #13#10#13#10;
end;

procedure TwndMSAAV.ThDone(Sender: TObject);
begin
  inc(iTHCnt);
	if not bTer then
	begin
  	if mnutvBoth.Checked then
		begin
			ShowText4UIA;
			TabSheet2.TabVisible := True;
		end;

		if (acShowTip.Checked) then
		begin
			ShowTipWnd;
		end;


    if Assigned(LoopNode) then
    begin
    	LoopNode.Text := 'Possibility of infinite loop(300 or more nested structure)';

  	end;
    GetNaviState;

	end
  else
  begin
    treeview1.Items.Clear;
    TBList.Clear;
    uTBList.Clear;
  end;

end;

procedure  TwndMSAAV.WMDPIChanged(var Message: TMessage);
begin
  scaleX := Message.WParamLo / DefX;//96.0;
  scaleY := Message.WParamHi / DefY;//96.0;
  if (not bFirstTime) and (cDPI <> Message.WParamLo) then
  begin

    cDPI := Message.WParamLo;
  	SizeChange;
  end;

end;
end.
