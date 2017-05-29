unit Thread;

interface

uses
  Classes, windows, SysUtils, Forms, ActiveX, Variants, ComCtrls, dialogs, WinAPI.Oleacc, iAccessible2Lib_tlb, MSHTML_tlb,
  UIAutomationClient_TLB, ISimpleDOM, IntList;

type

  TreeThread = class(TThread)
  private
    { Private êÈåæ }
    bReflex: boolean;
    iLvl: integer;
    UIAuto       : IUIAutomation;
    UIEle: IUIAutomationElement;
    iID: integer;
    procedure ReflexACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
  protected
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, sNode, RootNode: TTreeNode;
    NodeCap, None: string;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(IA: IAccessible; PA: IAccessible; AWnd: HWND; Mode: Integer; NoneText: string; CreateSuspended: boolean = True; Reflex: boolean = True; pNode: TTreenode = nil; accID: integer = 0); virtual;
  end;

  TreeThread4UIA = class(TThread)
  private
    { Private êÈåæ }
    bReflex: boolean;
    iLvl: integer;
    UIAuto       : IUIAutomation;
    UIEle, UIpEle: IUIAutomationElement;
    iVW: IUIAutomationTreeWalker;
    procedure ReflexACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement);
    function Get_RoleText(cEle: IUIAutomationElement): string;
  protected
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, sNode, RootNode: TTreeNode;
    NodeCap, None: string;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(UIA: IUIAutomation; iEle, pEle: IUIAutomationElement; NoneText: string; CreateSuspended: boolean = True; Reflex: boolean = True; pNode: TTreenode = nil); virtual;
  end;

implementation
uses
    frmMSAAV;


{ TreeThread }
constructor TreeThread.Create(IA: IAccessible; PA: IAccessible; AWnd: HWND; Mode: Integer; NoneText: string; CreateSuspended: boolean = True; Reflex: boolean = True; pNode: TTreenode = nil; accID: integer = 0);
begin
    IAcc := IA;
    pac := PA;
    RootNode := pNode;
    bReflex := Reflex;
    iLvl := 0;
    Imode := Mode;
    None := NoneText;
    iID := accID;

    inherited Create(CreateSuspended);
    FreeOnTerminate :=  false;
end;

procedure TreeThread.Execute;
var
	GetAcc: TThreadProcedure;
	hr: HResult;
begin
	GetAcc := procedure
  begin
    UIAuto.ElementFromIAccessible(iAcc, iID, UIEle);
  end;
  try
    CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
    frmMSAAV.bTer := false;
    hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER,
      IID_IUIAutomation, UIAuto);

    if (UIAuto = nil) or (hr <> S_OK) then
    begin
      Exit;

    end;

    try
      sNode := nil;
      Synchronize(GetAcc);
      ReflexACC(RootNode, pac, 0);

    except
    end;

    if Terminated then
    begin
      frmMSAAV.bTer := True;
      sNode := nil;
    end;

  finally
    CoUninitialize;
  end;
end;


function VarHaveValue(v: variant): boolean;
begin
    result := true;
    if VarType(v) = varEmpty then result := false;
    if VarIsEmpty(v) then result := false;
    if VarIsClear(v) then result := false;

end;



procedure TreeThread.ReflexACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
var
    cAcc, tAcc: iAccessible;
    iChild, i, iCH, iSame: integer;
    iObtain: plongint;
    cNode: TTreeNode;
    oc: OleVariant;
    iDis: iDispatch;
    TD: PTreeData;
    AddTreeItem, {SName, }SCnt, SCld: TThreadProcedure;
    Comp1: IUIAutomationElement;
    hr: HResult;
    aChildren   : array of TVariantArg;
begin
  if Terminated then
    Exit;
  AddTreeItem := procedure
  begin
      New(TD);
      TD^.Acc := cAcc;
      TD^.UIEle := nil;
      TD^.iID := iCH;
      pNode := cNode;
      rNode := wndMSAAV.TreeView1.Items.AddChildObject(pNode, '', Pointer(TD));
      TBList.Add(integer(rNode.ItemId));
      if sNode = nil then
      begin
        if SUCCEEDED(UIAuto.ElementFromIAccessible(cAcc, iCH, Comp1)) then
        begin
          UIAuto.CompareElements(UIEle, Comp1, iSame);
          if iSame <> 0 then
          begin
            sNode := rNode;
            sNode.Expanded := True;
            // wndMSAAV.TreeView1.SetFocus;
            wndMSAAV.TreeView1.TopItem := sNode;

            sNode.Selected := True;
          end;
        end;
      end;
  end;

  SCnt := procedure
  begin
    cAcc.Get_accChildCount(iChild);
  end;

  SCld := procedure
  var
  	i: integer;
  begin
  	cAcc.Get_accChildCount(iChild);
    SetLength(aChildren, iChild);
    for i := 0 to iChild - 1 do
    begin
      VariantInit(OleVariant(aChildren[i]));
    end;
  	if AccessibleChildren(cAcc, 0, iChild, @aChildren[0], iObtain) = S_OK then
    begin
    	for i := 0 to integer(iobtain) - 1 do
      begin
      	if Terminated then
      		Break;
      	tAcc := nil;
      	if aChildren[i].vt = VT_DISPATCH then
        begin
        	hr := IDispatch(aChildren[i].pdispVal).QueryInterface(IID_IACCESSIBLE, tAcc);
          if Assigned(tAcc) and (hr = S_OK) then
          begin
          	ReflexACC(cNode, tAcc, 0);
          end;
        end
        else if aChildren[i].vt = VT_I4 then
        begin
        	iCH := aChildren[i].lVal;
          Synchronize(AddTreeItem);
        end;
      end;
    end;

  end;

  cNode := ParentNode;
  cAcc := ParentAcc;
  iCH := iID;

  Synchronize(AddTreeItem);

  cNode := rNode;
  Synchronize(SCld);
end;

constructor TreeThread4UIA.Create(UIA: IUIAutomation; iEle, pEle: IUIAutomationElement; NoneText: string; CreateSuspended: boolean = True; Reflex: boolean = True; pNode: TTreenode = nil);
begin
    UIEle := iEle;
    UIpEle := pEle;
    RootNode := pNode;
    bReflex := Reflex;
    iLvl := 0;
    UIAuto := UIA;
    None := NoneText;


    inherited Create(CreateSuspended);
    FreeOnTerminate :=  false;
end;

procedure TreeThread4UIA.Execute;
begin
    try
        frmMSAAV.bTer := false;

        CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
    		if (not assigned(UIAuto)) or (not assigned(UIEle)) then
    		begin
        	Exit;

    		end;
        UIAuto.Get_ControlViewWalker(iVW);
        if not Assigned(iVW) then
          Exit;
        try
          sNode := nil;
          ReflexACC(RootNode, UIpEle);

        except
        end;

        if Terminated then
        begin
            frmMSAAV.bTer := True;
            snode := nil;
        end;

    finally
        CoUninitialize;
    end;
end;


{function VarHaveValue(v: variant): boolean;
begin
    result := true;
    if VarType(v) = varEmpty then result := false;
    if VarIsEmpty(v) then result := false;
    if VarIsClear(v) then result := false;

end;    }

function TreeThread4UIA.Get_RoleText(cEle: IUIAutomationElement): string;
var
    PC:PChar;
    ovValue: OleVariant;
begin

  try
    if SUCCEEDED(UIEle.GetCurrentPropertyValue(UIA_LegacyIAccessibleRolePropertyId, ovValue)) then
    begin
      if VarHaveValue(ovValue) then
      begin
        if VarIsNumeric(ovValue) then
        begin
          PC := StrAlloc(255);
          GetRoleTextW(ovValue, PC, StrBufSize(PC));
          result := PC;
          StrDispose(PC);
        end
        else if VarIsStr(ovValue) then
        begin
          result := VarToStr(ovValue);
        end;
      end;
    end;
  except
    result := None;
  end;
end;



procedure TreeThread4UIA.ReflexACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement);
var
  cEle, tEle, cldEle, psEle, SyncEle: IUIAutomationElement;
    iSame: integer;
    cNode: TTreeNode;
    ws: Widestring;
    Role: string;
    TD: PTreeData;
    SyncTree: TThreadProcedure;
    ovChild, ovValue: Olevariant;
    hr: HRESULT;

    procedure SyncName;
    begin
      if SUCCEEDED(UIEle.GetCurrentPropertyValue(UIA_LegacyIAccessibleNamePropertyId	, ovValue)) then
      begin
        ws := VarToStr(ovValue);
        if ws = '' then ws := None;
        Role := Get_ROLETExt(cEle);
      end;
    end;


    procedure SyncFCld;
    begin
      cldEle := nil;
      hr := iVW.GetFirstChildElement(cEle, cldEle);

    end;

    procedure SyncCld;
    begin
      tEle := nil;
      hr := iVW.GetFirstChildElement(cldEle, tEle);

    end;

    procedure SyncNCld;
    begin
      cldEle := nil;
      hr := iVW.GetNextSiblingElement(psEle, cldEle);
    end;
begin
    if Terminated then Exit;
    SyncTree := procedure
    begin
        New(TD);
        TD^.uiEle := SyncEle;
        TD^.Acc := nil;
        TD^.iID := 0;
        pNode := cNode;
        rNode := wndMSAAV.TreeView1.Items.AddChildObject(pNode, '', Pointer(TD));

        TBList.Add(Integer(rNode.ItemId));
        if sNode = nil then
        begin
          	UIAuto.CompareElements(UIEle, SyncEle, iSame);
            if iSame <> 0 then
            begin
            	sNode := rnode;
              sNode.Expanded := true;
              wndMSAAV.TreeView1.SetFocus;
              wndMSAAV.TreeView1.TopItem := sNode;

              sNode.Selected := True;
            end;
        end;
    end;

    cNode := ParentNode;
    cEle := ParentEle;
    ovChild := 0;
    SyncEle := cEle;
    Synchronize(SyncTree);


    cNode := rNode;
    SyncFCld;
    while (Assigned(cldEle)) and (hr = S_OK) do
    begin
      if Terminated then Break;
      SyncCld;
      if (Assigned(tEle)) and (hr = S_OK) then
        ReflexACC(cNode, cldEle)
      else
      begin
        SyncEle := cldEle;
        Synchronize(SyncTree);
      end;
      psEle := cldEle;
      SyncNCld;

    end;
end;

end.
