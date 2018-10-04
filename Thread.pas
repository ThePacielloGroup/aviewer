unit Thread;

interface

uses
  Classes, windows, SysUtils, Forms, ActiveX, Variants, ComCtrls, dialogs, WinAPI.Oleacc, iAccessible2Lib_tlb, MSHTML_tlb,
  UIAutomationClient_TLB, ISimpleDOM, IntList, winapi.commctrl, VirtualTrees, UIA_TLB, System.Math;

type

  TreeThread = class(TThread)
  private
    { Private êÈåæ }

    bRecursive: boolean;
    UIAuto       : IUIAutomation;
    UIEle: IUIAutomationElement;
    ipaID: integer;
    procedure RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
  protected
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, RootNode: TTreeNode;
    NodeCap, None, cAccRole: string;
    cAccName: widestring;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(IA: IAccessible; PA: IAccessible; AWnd: HWND; Mode: Integer; NoneText: string; bCreateSuspended: boolean = True; bRecur : boolean = True; pNode: TTreenode = nil; accID: integer = 0); virtual;
  end;

  TreeThread4UIA = class(TThread)
  private
    { Private êÈåæ }
    bRecursive: boolean;
    iLvl: integer;
    UIAuto       : IUIAutomation;
    UIEle, UIpEle: IUIAutomationElement;
    procedure RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
    procedure ExpandNode;
  protected
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, sNode, RootNode: TTreeNode;
    NodeCap, None: string;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(UIA: IUIAutomation; iEle, pEle: IUIAutomationElement; NoneText: string; bCreateSuspended: boolean = True; bRecur: boolean = True; pNode: TTreenode = nil); virtual;
  end;


  HTMLThread = class(TThread)
  private
    { Private êÈåæ }
    procedure GetHTMLItem;
  protected
    RootNode: PVirtualNode;
    cEle: IHTMLElement;
    procedure Execute; override;
  public
    constructor Create(iEle: IHTMLElement; pNode: PVirtualNode = nil; bCreateSuspended: boolean = True); virtual;
  end;

implementation
uses
    frmMSAAV;

function VarHaveValue(v: variant): boolean;
begin
    result := true;
    if VarType(v) = varEmpty then result := false;
    if VarIsEmpty(v) then result := false;
    if VarIsClear(v) then result := false;

end;

{HTMLThread}
constructor HTMLThread.Create(iEle: IHTMLElement; pNode: PVirtualNode = nil; bCreateSuspended: boolean = True);
begin
    RootNode := pNode;
    cEle := iEle;
    FreeOnTerminate :=  false;
    inherited Create(bCreateSuspended);

end;


procedure HTMLThread.Execute;
begin
	try
  	CoInitializeEx(nil, COINIT_APARTMENTTHREADED);
		if Assigned(CEle) then
		begin
  		GetHTMLItem;
		end;
  finally
  	CoUninitialize;
  end;
end;

procedure HTMLThread.GetHTMLItem;
var
  ND: PNodeData;
  ResNode, pNode: PVirtualNode;
  SetNodeData, SetMemoText, GetTagName, GetCollection, GetAttrSpec, GetAttrs, TreeExpand: TThreadProcedure;
  bSpecify: wordbool;
	s, Text1, Text2, nName: string;
	i, ColLen: integer;
	sNode: PVirtualNode;
	iAttrCol: IHTMLATTRIBUTECOLLECTION;
	iDomAttr: IHTMLDOMATTRIBUTE;
	ovValue, nValue: OleVariant;
begin

	SetNodeData := procedure
  begin
  	ResNode := wndMSAAV.TreeList1.InsertNode(pNode, amAddChildLast , nil);
  	ND := wndMSAAV.TreeList1.GetNodeData(ResNode);
  	ND.Value1 := Text1;
  	ND.Value2 := Text2;
  	ND.Acc := nil;
  	ND.iID := 0;
  end;

  SetMemoText := procedure
  begin
  	HTMLs[2, 1] := CEle.outerHTML;
  	wndMSAAV.Memo1.Text :=  HTMLs[2, 0] + ':' + #13#10 + HTMLs[2, 1] + #13#10#13#10;

  end;

  TreeExpand := procedure
  begin
  	wndMSAAV.TreeList1.Expanded[RootNode] := True;
    if Assigned(sNode) then
    	wndMSAAV.TreeList1.Expanded[sNode] := True;
  end;

  GetTagName := procedure
  begin
  	HTMLs[0, 1] := CEle.tagName;
  end;

  GetCollection := procedure
  begin
  	iAttrCol := (CEle as IHTMLDOMNode).attributes as IHTMLATTRIBUTECOLLECTION;
    ColLen := iAttrCol.Length;
  end;

  GetAttrSpec := procedure
  begin
  	iDomAttr := iAttrCol.item(ovValue) as IHTMLDOMATTRIBUTE;
    bSpecify := iDomAttr.specified;
  end;

  GetAttrs := Procedure
  begin
    nName := iDomAttr.nodeName;
    nValue := iDomAttr.nodeValue;
  end;
  Synchronize(SetMemoText);
  Synchronize(GetTagName);
  pNode := RootNode;
  Text1 := HTMLs[0, 0];
  Text2 := HTMLs[0, 1];
  Synchronize(SetNodeData);

  Synchronize(GetCollection);

		sNode := nil;
		for i := 0 to ColLen - 1 do
		begin
    	if Terminated then
				break;
			sleep(1);
			ovValue := i;
      Synchronize(GetAttrSpec);
			if bSpecify then
			begin
      	Synchronize(GetAttrs);
				s := LowerCase(nName);
				if (s <> 'role') and (copy(s, 1, 4) <> 'aria') then
				begin
					try
						if sNode = nil then
						begin
            	pNode := RootNode;
              Text1 := HTMLs[1, 0];
              Text2 := '';
              Synchronize(SetNodeData);
							sNode := ResNode;
						end;

						if VarHaveValue(nValue) then
						begin
            	pNode := sNode;
              Text1 := nName;
              Text2 := VarToStr(nValue);
              Synchronize(SetNodeData);
						end
						else
						begin
            	pNode := sNode;
              Text1 := nName;
              Text2 := '';
              Synchronize(SetNodeData);
						end;
					except
						on E: Exception do
						begin
            	MessageDlg(E.Message , mtError, [mbOK], 0);
						end;
					end;
				end;
			end;
		end;


  Synchronize(TreeExpand);
end;


{ TreeThread }
constructor TreeThread.Create(IA: IAccessible; PA: IAccessible; AWnd: HWND; Mode: Integer; NoneText: string; bCreateSuspended: boolean = True; bRecur: boolean = True; pNode: TTreenode = nil; accID: integer = 0);
begin
    IAcc := IA;
    pac := PA;
    RootNode := pNode;
    bRecursive := bRecur;
    Imode := Mode;
    None := NoneText;
    ipaID := accID;
    FreeOnTerminate :=  false;
    inherited Create(bCreateSuspended);

end;

procedure TreeThread.Execute;
var
	GetAcc: TThreadProcedure;
	hr: HResult;
begin
	GetAcc := procedure
  begin
    UIAuto.ElementFromIAccessible(iAcc, ipaID, UIEle);
    frmMSAAV.compEle := UIEle;
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
			Synchronize(GetAcc);
			frmMSAAV.LoopNode := nil;
			RecurACC(RootNode, pac, 0);

			if Terminated then
			begin
				frmMSAAV.bTer := True;

			end;
		except
		end;
	finally
    CoUninitialize;
  end;
end;



procedure TreeThread.RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
var
    cAcc, tAcc, SyncAcc: iAccessible;
    iChild, iCH, i: integer;
    iObtain: plongint;
    cNode: TTreeNode;
    TD: PTreeData;
    AddTreeItem, {SName, }SCnt, SCld, GetCld: TThreadProcedure;
    hr, hr_SCld: HResult;
    aChildren   : array of TVariantArg;
begin

  AddTreeItem := procedure
  begin
      New(TD);
      TD^.Acc := SyncAcc;
      TD^.UIEle := nil;
      TD^.iID := iCH;
      pNode := cNode;
      rNode := wndMSAAV.TreeView1.Items.AddChildObject(pNode, '', Pointer(TD));
      if (rNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := ParentNode;
      end;
      TBList.Add(integer(rNode.ItemId));
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
    hr_SCld := AccessibleChildren(cAcc, 0, iChild, @aChildren[0], iObtain);
  end;

  GetCld := procedure
  begin
  	hr := IDispatch(aChildren[i].pdispVal)
				.QueryInterface(IID_IACCESSIBLE, tAcc);
  end;

  if Terminated then
    Exit;
  cNode := ParentNode;
  cAcc := ParentAcc;
  SyncAcc := cAcc;
  iCH := iID;
  if Assigned(cNode) and (cNode.Level >= 300) and Assigned(frmMSAAV.LoopNode) then
  	exit;
  Synchronize(AddTreeItem);

  cNode := rNode;
  Synchronize(SCld);
  if hr_SCld = 0 then
	begin
		for i := 0 to integer(iObtain) - 1 do
		begin
			if Terminated then
				break;
      sleep(1);
			if aChildren[i].vt = VT_DISPATCH then
			begin
				Synchronize(GetCld);
				if Assigned(tAcc) and (hr = S_OK) then
				begin
					RecurACC(cNode, tAcc, 0);
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

constructor TreeThread4UIA.Create(UIA: IUIAutomation; iEle, pEle: IUIAutomationElement; NoneText: string; bCreateSuspended: boolean = True; bRecur: boolean = True; pNode: TTreenode = nil);
begin
    UIEle := iEle;
    UIpEle := pEle;
    RootNode := pNode;
    bRecursive := bRecur;
    iLvl := 0;
    //UIAuto := UIA;
    None := NoneText;
    FreeOnTerminate :=  false;

    inherited Create(bCreateSuspended);

end;

procedure TreeThread4UIA.Execute;
var
	uiCondi  : IUIAutomationCondition;
  ov: OleVariant;
  hr: hresult;
begin
    try
        frmMSAAV.bTer := false;

        CoInitializeEx(nil, {COINIT_APARTMENTTHREADED}COINIT_MULTITHREADED);
        hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER, IID_IUIAutomation, UIAuto);
    		if (not assigned(UIAuto)) or (not assigned(UIEle)) then
    		begin
        	Exit;

    		end;
        iLvl := 0;
      	frmMSAAV.LoopNode := nil;

        try
          sNode := nil;
          TVariantArg(ov).vt := VT_BOOL;
      		TVariantArg(ov).vbool := True;
      		hr := UIAuto.CreatePropertyCondition(UIA_IsControlElementPropertyId, ov, uiCondi);
          if (hr = 0) and (Assigned(uiCondi)) then
          begin
          	RecurACC(RootNode, UIpEle, uiCondi);
          	Synchronize(ExpandNode);
          end;
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

procedure TreeThread4UIA.ExpandNode;
begin
	if Assigned(sNode) then
  begin
		sNode.Expanded := True;
  	sNode.Selected := True;
  end;
end;

{function VarHaveValue(v: variant): boolean;
begin
    result := true;
    if VarType(v) = varEmpty then result := false;
    if VarIsEmpty(v) then result := false;
    if VarIsClear(v) then result := false;

end;    }



procedure TreeThread4UIA.RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
var
	tEle, cldEle, SyncEle: IUIAutomationElement;
	iSame, iImgIdx: integer;
	cNode: TTreeNode;
	ws: widestring;
	TD: PTreeData;
	SyncTree: TThreadProcedure;
	ov: OleVariant;
	hr: HResult;
	PC: PChar;
	arElement: IUIAutomationElementArray;
	iLen, i: integer;
	Scope: TreeScope;
	procedure GetNodeText;
	begin

		if SUCCEEDED(SyncEle.Get_CurrentName(ws)) then
		begin
			ws := StringReplace(ws, #13, ' ', [rfReplaceAll]);
			ws := StringReplace(ws, #10, ' ', [rfReplaceAll]);

			if ws = '' then
				ws := None;
			SyncEle.GetCurrentPropertyValue(UIA_LegacyIAccessibleRolePropertyId, ov);
			if VarIsType(ov, VT_I4) and (TVarData(ov).VInteger <= 61) then
			begin
				iImgIdx := TVarData(ov).VInteger - 1;
				PC := StrAlloc(255);
				GetRoleTextW(ov, PC, StrBufSize(PC));
				ws := ws + ' - ' + PC;
				StrDispose(PC);

			end;
		end;
	end;

begin
	if Terminated then
		Exit;
	if not Assigned(ParentEle) then
		Exit;

	SyncTree := procedure
		begin
			New(TD);
			TD^.UIEle := SyncEle;
			TD^.Acc := nil;
			TD^.iID := 0;
      rNode := wndMSAAV.TreeView1.Items.AddChildObject(cNode, ws, Pointer(TD));
			rNode.ImageIndex := iImgIdx;
			rNode.ExpandedImageIndex := iImgIdx;
			rNode.SelectedIndex := iImgIdx;
			TBList.Add(integer(rNode.ItemId));
			if (rNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
			begin
				frmMSAAV.LoopNode := rNode;
			end;
			if sNode = nil then
			begin
				UIAuto.CompareElements(UIEle, SyncEle, iSame);
				if iSame <> 0 then
				begin
					sNode := rNode;

				end;
			end;
		end;
	cNode := ParentNode;
	SyncEle := ParentEle;

	GetNodeText;
	Synchronize(SyncTree);
	cNode := rNode;

	Scope := TreeScope_Children;
	ParentEle.FindAll(Scope, uiCondition, arElement);
	arElement.Get_Length(iLen);
	for i := 0 to iLen - 1 do
	begin
		hr := arElement.GetElement(i, cldEle);
		if (hr = 0) and (Assigned(cldEle)) then
		begin
			hr := cldEle.FindFirst(Scope, uiCondition, tEle);
			if (hr = 0) and (Assigned(tEle)) then
				RecurACC(cNode, cldEle, uiCondition)
			else
			begin
				SyncEle := cldEle;
				GetNodeText;
				Synchronize(SyncTree);
			end;
		end;
		if Terminated then
			break;
    sleep(1);
	end;

end;

end.
