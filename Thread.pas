unit Thread;

interface

uses
  Classes, windows, SysUtils, Forms, ActiveX, Variants, ComCtrls, dialogs, WinAPI.Oleacc, iAccessible2Lib_tlb, MSHTML_tlb,
  UIAutomationClient_TLB, ISimpleDOM, IntList, winapi.commctrl, VirtualTrees, UIA_TLB, System.Math, System.RegularExpressions, StrUtils;

type

  TreeThread = class(TThread)
  private
    { Private �錾 }

    bRecursive: boolean;
    ipaID: integer;
    procedure GetSelNode;
    procedure RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
    procedure ExpandNode;
    function Get_RoleText(Acc: IAccessible; Child: integer): string;
    function IsSameUIElement(ia1, ia2: IAccessible; iID1, iID2: integer): boolean;
    function AddTreeItem(AccItem: iAccessible; AccID: integer; ParentItem: TTreeNode; NewNode: boolean = true; DummyNode: boolean = false): TTreenode;
    procedure SetParentTree;
    procedure RecursiveGetTreeitemID(currentNode: TTreenode);
  protected
  	dMode, bAllEx: boolean;
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, sNode, RootNode, SelNode, SelPaNode: TTreeNode;
    None, cAccRole: string;
    cAccName: widestring;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(IA: IAccessible; PA: IAccessible; NoneText: string; bCreateSuspended: boolean = True; bRecur : boolean = True; pNode: TTreenode = nil; accID: integer = 0; DummyMode : boolean = false; AllEx : boolean = false); virtual;
  end;

  TreeThread4UIA = class(TThread)
  private
    { Private �錾 }
    iVW: IUIAutomationTreeWalker;
    bRecursive: boolean;
    iLvl: integer;
    UIAuto       : IUIAutomation;
    UIEle, UIpEle: IUIAutomationElement;
    procedure RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
    procedure ExpandNode;
    function AddTreeItem(iUIEle: IUIAutomationElement; ParentItem: TTreeNode; NewNode: boolean = true): TTreenode;
    procedure GetSelNode;
    procedure SetParentTree(uiCondition: IUIAutomationCondition);
    function IsSameUIElement(uEle1, uEle2: IUIAutomationElement): boolean;
    procedure RecursiveGetTreeitemID(currentNode: TTreenode);
  protected
    iAcc, pac, getcAcc, tAcc: IAccessible;
    pNode, rNode, sNode, RootNode, SelNode, SelPaNode: TTreeNode;
    NodeCap, None: string;
    iMode: Integer;
    RC: TRect;
    Wnd: hwnd;
    procedure Execute; override;
  public
    constructor Create(UIA: IUIAutomation; iEle, pEle: IUIAutomationElement; NoneText: string; bCreateSuspended: boolean = True; bRecur: boolean = True; pNode: TTreenode = nil); virtual;
  end;

  MSAAExTh = class(TThread)
  private
    { Private �錾 }
    function AddTreeItem(AccItem: iAccessible; AccID: integer; ParentItem: TTreeNode; NewNode: boolean = true; DummyNode: boolean = false): TTreenode;
    procedure RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
  protected
    procedure Execute; override;
  public
    constructor Create(bCreateSuspended: boolean = True); virtual;
  end;

  UIAExTh = class(TThread)
  private
    { Private �錾 }
    UIAuto       : IUIAutomation;
    function AddTreeItem(iUIEle: IUIAutomationElement; ParentItem: TTreeNode; NewNode: boolean = true): TTreenode;
    procedure RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
  protected
    procedure Execute; override;
  public
    constructor Create(bCreateSuspended: boolean = True); virtual;
  end;

  HTMLThread = class(TThread)
  private
    { Private �錾 }
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
  SetNodeData, SetMemoText, GetTagName, TreeExpand: TThreadProcedure;
	Text1, Text2: string;
  hAttrs: String;
	i: integer;
	sNode: PVirtualNode;

  match: TMatch;
  matches: TMatchCollection;
  iPos: integer;
  atname, atValue, q: string;
begin

	SetNodeData := procedure
  begin
  	//ResNode := wndMSAAV.TreeList1.InsertNode(pNode, amAddChildLast , nil);
    ResNode := wndMSAAV.TreeList1.AddChild(pNode, nil);
  	ND := ResNode.GetData;
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

  Synchronize(SetMemoText);
  Synchronize(GetTagName);
  pNode := RootNode;
  Text1 := HTMLs[0, 0];
  Text2 := HTMLs[0, 1];
  Synchronize(SetNodeData);

  sNode := nil;
  match := TRegEx.Match(HTMLs[2, 1], '<("[^"]*"|''[^'']*''|[^''">])*>');
  if match.Success then
  begin
  	matches := TRegEx.matches(match.Value, '(\S+)=[""'']?((?:.(?![""'']?\s+(?:\S+)=|[>""'']))+.)[""'']?');

    for i := 0 to matches.Count - 1 do
		begin
    	if sNode = nil then
			begin
				pNode := RootNode;
				Text1 := HTMLs[1, 0];
				Text2 := '';
				Synchronize(SetNodeData);
				sNode := ResNode;
			end;
			iPos := Pos('=', matches.Item[i].Value);
			atname := Copy(matches.Item[i].Value, 1, iPos - 1);
			atValue := Copy(matches.Item[i].Value, iPos + 1,
				Length(matches.Item[i].Value));
			q := Copy(atValue, 1, 1);
			if (q = '''') or (q = '"') then
			begin
				atValue := Copy(atValue, 2, Length(atValue));
				atValue := Copy(atValue, 1, Length(atValue) - 1);
			end;
      pNode := sNode;
      Text1 := atname;
      Text2 := atValue;
      hAttrs := hAttrs + Text1 + '=' + Text2 + #13#10;
      Synchronize(SetNodeData);

		end;
  end;

  if hAttrs = '' then
  	HTMLs[1, 1] := None
  else
  	HTMLs[1, 1] := hAttrs;

  Synchronize(TreeExpand);
end;

{ MSAA Expose Thread }

function MSAAExTh.AddTreeItem(AccItem: iAccessible; AccID: integer; ParentItem: TTreeNode; NewNode: boolean = true; DummyNode: boolean = false): TTreenode;
var
	AddItem: TThreadProcedure;
  TD: PTreeData;
  resNode: TTreeNode;
begin
  AddItem := procedure
  begin
      New(TD);
      TD^.Acc := AccItem;
      TD^.UIEle := nil;
      TD^.iID := AccID;
      TD^.dummy := DummyNode;
      if NewNode then
      begin
      	resNode := wndMSAAV.TreeView1.Items.AddChildObject(ParentItem, '', Pointer(TD));
      end
      else
      begin
      	ParentItem.Text := '';
        ParentItem.Data := Pointer(TD);
        resNode := ParentItem;
        wndMSAAV.TreeView1.OnAddition(wndMSAAV.TreeView1, resNode);
      end;
      TBList.Add(integer(resNode.ItemId));
      if (resNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := ParentItem;
      end;
  end;
  Synchronize(AddItem);
  Result := resNode;
end;

constructor MSAAExTh.Create(bCreateSuspended: boolean = True);
begin
	FreeOnTerminate :=  false;
	inherited Create(bCreateSuspended);
end;

procedure MSAAExTh.Execute;
var
	GetNodeD, SetPBar: TThreadProcedure;
	hr: HResult;
  i, cID, iTarg, iProg: integer;
  cAcc: iAccessible;
  cNode: TTreeNode;
  bDummy: boolean;
begin
		GetNodeD := procedure
		begin
			cAcc := nil;
      bDummy := False;
			cNode := wndMSAAV.TreeView1.Items.GetNode(HTreeItem(DList.Items[iTarg]));
			if (Assigned(cNode)) and Assigned(cNode.Data) then
			begin
				cAcc := TTreeData(cNode.Data^).Acc;
				cID := TTreeData(cNode.Data^).iID;
        bDummy := TTreeData(cNode.Data^).dummy;
        if (bDummy) and (cNode.HasChildren) and  (cNode.Item[0].Text = 'avwr_dummy') then
        	cNode.DeleteChildren;
			end;
		end;
    SetPBar := procedure
    begin
    	wndMSAAV.Taskbar1.ProgressValue := iProg;
    end;
	try
  	iProg := 0;
    Synchronize(SetPBar);
  	for i := 0 to DList.Count - 1 do
    begin
    	if Terminated then
      	break;
      iProg := Round((i / DList.Count) * 100);
    	Synchronize(SetPBar);
      iTarg := i;
      Synchronize(GetNodeD);
      if bDummy then
      begin
      	RecurACC(cNode, cAcc, cID);
        TTreeData(cNode.Data^).dummy := False;
      end;
    end;
	finally
  end;
end;

procedure MSAAExTh.RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
var
    cAcc, tAcc: iAccessible;
    iChild, iCH, i: integer;
    iObtain: plongint;
    cNode, Recurnode: TTreeNode;
    SCnt, SCld, GetCld: TThreadProcedure;
    hr, hr_SCld: HResult;
    aChildren   : array of TVariantArg;
begin



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
  if Assigned(cNode) and (cNode.Level >= 300) and Assigned(frmMSAAV.LoopNode) then
  	exit;


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
        	RecurNode := AddTreeItem(tAcc, 0, cNode, true, true);

					RecurACC(RecurNode, tAcc, 0);
				end;
			end
			else if aChildren[i].vt = VT_I4 then
			begin
				iCH := aChildren[i].lVal;
				AddTreeItem(cAcc, iCH, cNode, true, true);
			end;
		end;
	end;
end;



{ TreeThread }
constructor TreeThread.Create(IA: IAccessible; PA: IAccessible; NoneText: string; bCreateSuspended: boolean = True; bRecur: boolean = True; pNode: TTreenode = nil; accID: integer = 0; DummyMode: boolean = false; AllEx : boolean = false);
begin
    IAcc := IA;
    pac := PA;
    RootNode := pNode;
    bRecursive := bRecur;
    None := NoneText;
    ipaID := accID;
    DMode := Dummymode;
    bAllEx := AllEx;
    FreeOnTerminate :=  false;
    inherited Create(bCreateSuspended);

end;

function TreeThread.Get_RoleText(Acc: IAccessible; Child: integer): string;
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

    end;
end;

procedure TreeThread.Execute;
var
	GetNodeD: TThreadProcedure;
	hr: HResult;
  i, cID, iTarg: integer;
  cAcc: iAccessible;
  cNode: TTreeNode;
begin

  GetNodeD := procedure
		begin
			cAcc := nil;
			cNode := wndMSAAV.TreeView1.Items.GetNode(HTreeItem(DList.Items[iTarg]));
			if (Assigned(cNode)) then
			begin
				cAcc := TTreeData(cNode.Data^).Acc;
				cID := TTreeData(cNode.Data^).iID;
				wndMSAAV.TreeView1.Items.AddChild(cNode, 'avwr_dummy');
			end;
		end;

	try
		SelNode := nil;

		try
			sNode := nil;
			frmMSAAV.LoopNode := nil;
			if not dMode then
			begin
				GetSelNode;
				if Assigned(SelNode) then
				begin
					SetParentTree;

					if bRecursive then
						RecursiveGetTreeitemID(SelPaNode);

					for i := 0 to DList.Count - 1 do
					begin
						if Terminated then
							break;
						iTarg := i;
						Synchronize(GetNodeD);

						sleep(1);
					end;
					Synchronize(ExpandNode);
				end;
			end
			else
      begin
      	if Assigned(iAcc) then
        begin
          if bAllEx then
						RecurACC(RootNode, iAcc, ipaID)
          else
          begin
          	SelNode := pNode;
          end;

        end;
      end;
		except
		end;
	finally
	end;
end;

procedure TreeThread.RecursiveGetTreeitemID(currentNode: TTreenode);
var
	i: integer;
  cNode: TTreeNode;
  GetValue: TThreadProcedure;
  bDummy: boolean;
begin
  GetValue := procedure
  begin
  	bDummy := False;
    if Assigned(cNode.Data) then
  		bDummy := TTreeData(cNode.Data^).dummy;
  end;
	if not Assigned(currentNode) then
  	Exit;
  cNode := currentNode;
  Synchronize(GetValue);
  if not bDummy then
  		TBList.Insert(0, integer(currentNode.ItemId))
  	else
  		DList.Insert(0, integer(currentNode.ItemId));
  if currentNode.HasChildren then
  begin
  	for i := 0 to currentNode.Count - 1 do
    begin
    	if Terminated then
      	break;
    	if currentNode.Item[i].HasChildren then
      	RecursiveGetTreeitemID(currentNode.Item[i])
      else
      begin
      	cNode := currentNode.Item[i];
  			Synchronize(GetValue);
        if not bDummy then
  				TBList.Insert(0, integer(currentNode.Item[i].ItemId))
  			else
  				DList.Insert(0, integer(currentNode.Item[i].ItemId));
      end;
      Sleep(1);
    end;
  end;

end;

function TreeThread.IsSameUIElement(ia1, ia2: IAccessible; iID1, iID2: integer): boolean;
var
    hr: hresult;
    cRC, cRC2: TRect;
    wsValue1, wsValue2: widestring;
    sRole1, sRole2: string;
    GetValue: TThreadProcedure;
begin
	GetValue := procedure
    begin
    	ia1.Get_accValue(iID1, wsValue1);
    	sRole1 := Get_RoleText(ia1, iID1);
    	hr := ia1.accLocation(cRC.Left, cRC.Top, cRC.Right, cRC.Bottom, iID1);

      ia2.Get_accValue(iID2, wsValue2);
      sRole2 := Get_RoleText(ia2, iID2);
      hr := ia2.accLocation(cRC2.Left, cRC2.Top, cRC2.Right, cRC2.Bottom, iID2);
    end;
	Result := false;
	if Assigned(ia1) and Assigned(ia2) then
	begin

		Synchronize(GetValue);

		if (wsValue1 = wsValue2) and (sRole1 = sRole2) and (cRC.Left = cRC2.Left)
			and (cRC.Top = cRC2.Top) and (cRC.Right = cRC2.Right) and
			(cRC.Bottom = cRC2.Bottom) then
			result := true;

	end;
end;


procedure TreeThread.ExpandNode;
begin
  Selnode.Expanded := True;
  Selnode.Selected := True;
end;

function TreeThread.AddTreeItem(AccItem: iAccessible; AccID: integer; ParentItem: TTreeNode; NewNode: boolean = true; DummyNode: boolean = false): TTreenode;
var
	AddItem: TThreadProcedure;
  TD: PTreeData;
  resNode: TTreeNode;
begin
  AddItem := procedure
  begin
      New(TD);
      TD^.Acc := AccItem;
      TD^.UIEle := nil;
      TD^.iID := AccID;
      TD^.dummy := DummyNode;
      if NewNode then
      begin
      	resNode := wndMSAAV.TreeView1.Items.AddChildObject(ParentItem, '', Pointer(TD));
      end
      else
      begin
      	ParentItem.Text := '';
        ParentItem.Data := Pointer(TD);
        resNode := ParentItem;
        wndMSAAV.TreeView1.OnAddition(wndMSAAV.TreeView1, resNode);
      end;
      TBList.Add(integer(resNode.ItemId));
      if (resNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := ParentItem;
      end;
  end;
  Synchronize(AddItem);
  Result := resNode;
end;

procedure TreeThread.GetSelNode;
begin
  SelNode := AddTreeItem(iAcc, ipaID, RootNode);
  Selnode.Expand(True);
  Selnode.Selected := True;
end;

procedure TreeThread.SetParentTree;
var
    cAcc, tAcc, SyncAcc, pAcc, compAcc: iAccessible;
    iChild, iCH, i, itarg, iComp, dChild: integer;
    iObtain: plongint;
    cNode,dNode, SSelNode: TTreeNode;
    TD: PTreeData;
    SCnt, SCld, GetCld, AddDummyNode, GetPaAcc, NodeMove, paExpand: TThreadProcedure;
    hr, hr_SCld: HResult;
    aChildren   : array of TVariantArg;
    bFSame, bAddDL: boolean;
    iDis: IDispatch;
begin



  AddDummyNode := procedure
  begin
  	New(TD);
      TD^.Acc := tAcc;
      TD^.UIEle := nil;
      TD^.iID := iCH;
      TD^.dummy := true;
      pNode := cNode;
      dNode := wndMSAAV.TreeView1.Items.AddChildObject(pNode, '', Pointer(TD));

      if (dNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := dNode;
      end;
      TBList.Add(integer(dNode.ItemId));
      if bAddDL then
      begin
      	tAcc.Get_accChildCount(dChild);
      	if (dChild > 0)  then
      	begin
					DList.Add(integer(dNode.ItemId));

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
    hr_SCld := AccessibleChildren(cAcc, 0, iChild, @aChildren[0], iObtain);
  end;

  GetCld := procedure
  begin
  	tAcc := nil;
  	hr := IDispatch(aChildren[itarg].pdispVal)
				.QueryInterface(IID_IACCESSIBLE, tAcc);
  end;

  GetPaAcc := procedure
  begin
  	pAcc := nil;
    if iCH = 0 then
    begin
  		hr := cAcc.Get_accParent(iDis);
    	if (hr = S_OK) and (Assigned(iDis)) then
    		hr := iDis.QueryInterface(IID_IACCESSIBLE, pAcc);
    end
    else
    begin
    	hr := S_OK;
      pAcc := cAcc;
    end;
  end;

  NodeMove := procedure
  begin
  	SSelNode.MoveTo(cNode, naAddChild);
    SSelNode := cNode;
    if SSelNode.Level = 0 then
    	SelPaNode := SSelNode.Item[iTarg];
  end;

  if Terminated then
    Exit;

  cAcc := iAcc;
  SyncAcc := cAcc;
  pNode := SelNode;
  cNode := pNode;
  iCH := ipaID;
  bAddDL := false;
  if ipaID = 0 then
	begin

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
        	itarg := i;
					Synchronize(GetCld);
					if Assigned(tAcc) and (hr = S_OK) then
					begin
						iCH := 0;
            if not bRecursive then bAddDL := true;
						Synchronize(AddDummyNode);
					end;
				end
        else
        begin
        	iCH := aChildren[i].lVal;
					rNode := AddTreeItem(SyncAcc, iCH, pNode, true, false);
        end;
        Sleep(1);
			end;
		end;
	end;
  Synchronize(ExpandNode);
	if bRecursive then
  begin
    bAddDL := true;
    Synchronize(GetPaAcc);
    SSelNode := Selnode;
    compAcc := iAcc;
    iComp := ipaID;
    while ((hr = S_OK) and Assigned(pAcc)) do
    begin
      if Terminated then
      	break;
    	SyncAcc := pAcc;
      cAcc := pAcc;
      iCH := 0;
      cNode := nil;
      rNode := AddTreeItem(SyncAcc, iCH, cNode);

      bFSame := False;
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
          	iCH := 0;
            itarg := i;
						Synchronize(GetCld);

          end
          else
          begin
          	iCH := aChildren[i].lVal;
            tAcc := pAcc;
          end;

          if (not bFSame) and (IsSameUIElement(compAcc, tAcc, iComp, iCH)) then
          begin
          	itarg := i;
          	Synchronize(NodeMove );
            bFSame := True;
            compAcc := pAcc;
      			iComp := 0;

          end
          else
          begin

          	if aChildren[i].vt = VT_DISPATCH then
						begin
							iCH := 0;
							Synchronize(AddDummyNode);

						end
						else
						begin
							iCH := aChildren[i].lVal;
							rNode := AddTreeItem(SyncAcc, iCH, cNode, true, false);
						end;
          end;
          sleep(1);
          rNode.Expand(true);
				end;   //for end
			end;

      if (IsSameUIElement(pAcc, pac, 0, 0)) then  break;
      iCH := 0;
      Synchronize(GetPaAcc);
      sleep(1);
		end; //while end
  end;
end;


procedure TreeThread.RecurACC(ParentNode: TTreeNode; ParentAcc: iAccessible; iID: integer);
var
    cAcc, tAcc: iAccessible;
    iChild, iCH, i: integer;
    iObtain: plongint;
    cNode, Recurnode: TTreeNode;
    SCnt, SCld, GetCld: TThreadProcedure;
    hr, hr_SCld: HResult;
    aChildren   : array of TVariantArg;
begin



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
  if Assigned(cNode) and (cNode.Level >= 300) and Assigned(frmMSAAV.LoopNode) then
  	exit;


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
        	RecurNode := AddTreeItem(tAcc, 0, cNode, true, true);

					RecurACC(RecurNode, tAcc, 0);
				end;
			end
			else if aChildren[i].vt = VT_I4 then
			begin
				iCH := aChildren[i].lVal;
				AddTreeItem(cAcc, iCH, cNode, true, true);
			end;
		end;
	end;
end;

constructor UIAExTh.Create(bCreateSuspended: boolean = True);
begin
    FreeOnTerminate :=  false;
    inherited Create(bCreateSuspended);

end;

procedure UIAExTh.Execute;
var
	uiCondi  : IUIAutomationCondition;
  ov: OleVariant;
  hr: hresult;
  i, iTarg, iProg: integer;
  GetNodeD, SetPBar: TThreadProcedure;
  cEle: IUIAUTOMATIONELEMENT;
  cNode: TTreeNode;
  bDummy: boolean;
begin
	GetNodeD := procedure
  begin
  	cEle := nil;
    cNode := wndMSAAV.tbUIA.Items.GetNode(HTreeItem(uDList.Items[iTarg]));
    if (Assigned(cNode)) and (Assigned(TTreeData(cNode.Data^).uiEle)) then
    begin
    	cEle := TTreeData(cNode.Data^).uiEle;
      bDummy := TTreeData(cNode.Data^).dummy;
      if (bDummy) and (cNode.HasChildren) and  (cNode.Item[0].Text = 'avwr_dummy') then
        	cNode.DeleteChildren;
    end;
  end;

  SetPBar := procedure
  begin
  	wndMSAAV.Taskbar1.ProgressValue := iProg;
  end;
  try


  	CoInitializeEx(nil, { COINIT_APARTMENTTHREADED } COINIT_MULTITHREADED);
		hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER,
			IID_IUIAutomation, UIAuto);

		TVariantArg(ov).vt := VT_BOOL;
		TVariantArg(ov).vbool := true;
		hr := UIAuto.CreatePropertyCondition(UIA_IsControlElementPropertyId,
			ov, uiCondi);
		if (hr = 0) and (Assigned(uiCondi)) then
		begin
      iProg := 0;
    	Synchronize(SetPBar);
			for i := 0 to uDList.Count - 1 do
			begin
				if Terminated then
					break;
        iProg := Round((i / uDList.Count) * 100);
    		Synchronize(SetPBar);
				iTarg := i;
				Synchronize(GetNodeD);
				if Assigned(cEle) and (bDummy) then
        begin

        	RecurACC(cNode, cEle, uiCondi);

        end;
				sleep(1);
			end;
		end;
	finally
		CoUninitialize;
	end;
end;

function UIAExTh.AddTreeItem(iUIEle: IUIAutomationElement; ParentItem: TTreeNode; NewNode: boolean = true): TTreenode;
var
	AddItem: TThreadProcedure;
  TD: PTreeData;
  resNode: TTreeNode;
begin
	AddItem := procedure
  begin
      New(TD);
      TD^.Acc := nil;
      TD^.UIEle := iUIEle;
      TD^.iID := 0;
      TD^.dummy := false;
      if NewNode then
      begin
      	resNode := wndMSAAV.tbUIA.Items.AddChildObject(ParentItem, '', Pointer(TD));
      end
      else
      begin
      	ParentItem.Text := '';
        ParentItem.Data := Pointer(TD);
        resNode := ParentItem;
        wndMSAAV.tbUIA.OnAddition(wndMSAAV.tbUIA, resNode);
      end;
      uTBList.Add(integer(resNode.ItemId));
      if (resNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := ParentItem;
      end;
  end;
  Synchronize(AddItem);
  Result := resNode;

end;

procedure UIAExTh.RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
var
	tEle, SyncEle, fcldEle: IUIAutomationElement;
	iImgIdx: integer;
	cNode, Recurnode: TTreeNode;
	ws: widestring;
	SCld, GetCld, GetFCld: TThreadProcedure;
	ov: OleVariant;
	hr, hr_SCld: HResult;
	PC: PChar;
	arElement: IUIAutomationElementArray;
	iLen, i, iTarg: integer;
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

  SCld := procedure
  begin
  	Scope := TreeScope_Children;
		hr_SCld := SyncEle.FindAll(Scope, uiCondition, arElement);
		arElement.Get_Length(iLen);

  end;

  GetCld := procedure
  begin
  	tEle := nil;
    hr := arElement.GetElement(itarg, tEle);
  end;


  GetFCld := procedure
  begin
  	fcldEle := nil;
    Scope := TreeScope_Children;
    hr := tEle.FindFirst(Scope, uiCondition, fcldEle);
  end;

  cNode := ParentNode;
	SyncEle := ParentEle;
  if Assigned(cNode) and (cNode.Level >= 300) and Assigned(frmMSAAV.LoopNode) then
  	exit;


  if hr = S_OK then
  begin
    Synchronize(SCld);
    for i := 0 to iLen - 1 do
    begin
    	if Terminated then
				break;
			sleep(1);
			iTarg := i;
			Synchronize(GetCld);
			Synchronize(GetFCld);

			if Assigned(fcldEle) then
			begin
				Recurnode := AddTreeItem(tEle, cNode);
				RecurACC(Recurnode, tEle, uiCondition);
			end
			else
			begin
				AddTreeItem(tEle, cNode);
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
  i, iTarg: integer;
  GetNodeD, GetTWalker: TThreadProcedure;
  cEle: IUIAUTOMATIONELEMENT;
  cNode: TTreeNode;

begin
	GetNodeD := procedure
  begin
  	cEle := nil;
    cNode := wndMSAAV.tbUIA.Items.GetNode(HTreeItem(uDList.Items[iTarg]));
    if (Assigned(cNode)) and (Assigned(TTreeData(cNode.Data^).uiEle)) then
    begin
    	cEle := TTreeData(cNode.Data^).uiEle;
      wndMSAAV.tbUIA.Items.AddChild(cNode, 'avwr_dummy');
    end;
  end;

  GetTWalker := procedure
  begin
  	UIAuto.Get_ControlViewWalker(iVW);
  end;
    try


        CoInitializeEx(nil, {COINIT_APARTMENTTHREADED}COINIT_MULTITHREADED);
        hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER, IID_IUIAutomation, UIAuto);
    		if (not assigned(UIAuto)) or (not assigned(UIEle)) then
    		begin
        	Exit;

    		end;
        iLvl := 0;
      	frmMSAAV.LoopNode := nil;
        Synchronize(GetTWalker);
        GetSelNode;
      	if Assigned(SelNode) then
				begin
					try
						sNode := nil;
						TVariantArg(ov).vt := VT_BOOL;
						TVariantArg(ov).vbool := true;
						hr := UIAuto.CreatePropertyCondition(UIA_IsControlElementPropertyId,
							ov, uiCondi);
						if (hr = 0) and (Assigned(uiCondi)) then
						begin
							SetParentTree(uiCondi);

							if bRecursive then
								RecursiveGetTreeitemID(SelPaNode);

							for i := 0 to uDList.Count - 1 do
							begin
								if Terminated then
									break;
								iTarg := i;
								Synchronize(GetNodeD);
								{if Assigned(cEle) then
								begin

									RecurACC(cNode, cEle, uiCondi);

								end; }
								sleep(1);
							end;
							Synchronize(ExpandNode);
						end;
					except
					end;

					if Terminated then
					begin

						sNode := nil;
					end;
				end;
			finally
        CoUninitialize;
    end;
end;

procedure TreeThread4UIA.GetSelNode;
begin
  SelNode := AddTreeItem(UIEle, RootNode);
  Selnode.Expanded := True;
  Selnode.Selected := True;
end;


procedure TreeThread4UIA.ExpandNode;
begin
	if Assigned(sNode) then
  begin
		sNode.Expanded := True;
  	sNode.Selected := True;
  end;
end;

function TreeThread4UIA.AddTreeItem(iUIEle: IUIAutomationElement; ParentItem: TTreeNode; NewNode: boolean = true): TTreenode;
var
	AddItem: TThreadProcedure;
  TD: PTreeData;
  resNode: TTreeNode;
begin
	AddItem := procedure
  begin
      New(TD);
      TD^.Acc := nil;
      TD^.UIEle := iUIEle;
      TD^.iID := 0;
      TD^.dummy := false;
      if NewNode then
      begin
      	resNode := wndMSAAV.tbUIA.Items.AddChildObject(ParentItem, '', Pointer(TD));
      end
      else
      begin
      	ParentItem.Text := '';
        ParentItem.Data := Pointer(TD);
        resNode := ParentItem;
        wndMSAAV.tbUIA.OnAddition(wndMSAAV.tbUIA, resNode);
      end;
      uTBList.Add(integer(resNode.ItemId));
      if (resNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := ParentItem;
      end;
  end;
  Synchronize(AddItem);
  Result := resNode;

end;

function TreeThread4UIA.IsSameUIElement(uEle1, uEle2: IUIAutomationElement): boolean;
var
	iRole1, iRole2: integer;
  ws1, ws2: widestring;
  ovRole: OleVariant;
begin
	Result := False;
  uEle1.Get_CurrentName(ws1);
  uEle2.Get_CurrentName(ws2);
  iRole1 := 0;
  iRole2 := 0;
  if SUCCEEDED(uEle1.GetCurrentPropertyValue
		(UIA_LegacyIAccessibleRolePropertyId, ovRole)) then
	begin
		if VarHaveValue(ovRole) then
		begin
			if VarIsType(ovRole, VT_I4) then
				iRole1 := TVarData(ovRole).VInteger;
		end;
	end;
  if SUCCEEDED(uEle2.GetCurrentPropertyValue
		(UIA_LegacyIAccessibleRolePropertyId, ovRole)) then
	begin
		if VarHaveValue(ovRole) then
		begin
			if VarIsType(ovRole, VT_I4) then
				iRole2 := TVarData(ovRole).VInteger;
		end;
	end;

  if (ws1 = ws2) and (iRole1 = iRole2) then
  	Result := True;
	{hr := UIAuto.CompareElements(uEle1, uEle2, iSame);
  if SUCCEEDED(hr) and (iSame <> 0) then
  	Result := True; }

end;

procedure TreeThread4UIA.RecursiveGetTreeitemID(currentNode: TTreenode);
var
	i: integer;
  cNode: TTreeNode;
  GetValue: TThreadProcedure;
  bDummy: boolean;
begin
  GetValue := procedure
  begin
  	bDummy := False;
    if Assigned(cNode.Data) then
  		bDummy := TTreeData(cNode.Data^).dummy;
  end;
	if not Assigned(currentNode) then
  	Exit;
  cNode := currentNode;
  Synchronize(GetValue);
  if not bDummy then
  		uTBList.Insert(0, integer(currentNode.ItemId))
  	else
  		uDList.Insert(0, integer(currentNode.ItemId));
  if currentNode.HasChildren then
  begin
  	for i := 0 to currentNode.Count - 1 do
    begin
    	if Terminated then
      	break;
    	if currentNode.Item[i].HasChildren then
      	RecursiveGetTreeitemID(currentNode.Item[i])
      else
      begin
      	cNode := currentNode.Item[i];
  			Synchronize(GetValue);
        if not bDummy then
  				uTBList.Insert(0, integer(currentNode.Item[i].ItemId))
  			else
  				uDList.Insert(0, integer(currentNode.Item[i].ItemId));
      end;
      Sleep(1);
    end;
  end;

end;

procedure TreeThread4UIA.SetParentTree(uiCondition: IUIAutomationCondition);
var
	cEle, tEle, SyncEle, pEle, compEle, fcldEle: IUIAutomationElement;
	arElement: IUIAutomationElementArray;
	iLen: integer;
	Scope: TreeScope;


    i, itarg: integer;
    cNode,dNode, SSelNode: TTreeNode;
    TD: PTreeData;
    SCld, GetCld, AddDummyNode, GetPaAcc, NodeMove, GetFCld: TThreadProcedure;
    hr, hr_SCld: HResult;
    bFSame: boolean;


begin



  AddDummyNode := procedure
  begin
  	New(TD);
      TD^.Acc := nil;
      TD^.UIEle := tEle;
      TD^.iID := 0;
      TD^.dummy := true;
      pNode := cNode;
      dNode := wndMSAAV.tbUIA.Items.AddChildObject(pNode, '', Pointer(TD));
      if (dNode.Level >= 300) and not Assigned(frmMSAAV.LoopNode) then
      begin
      	frmMSAAV.Loopnode := dNode;
      end;
      uTBList.Add(integer(dNode.ItemId));
      uDList.Add(integer(dNode.ItemId));
  end;

  SCld := procedure
  begin
  	Scope := TreeScope_Children;
		hr_SCld := cEle.FindAll(Scope, uiCondition, arElement);
		arElement.Get_Length(iLen);

  end;

  GetCld := procedure
  begin
  	tEle := nil;
    hr := arElement.GetElement(itarg, tEle);
  end;

  GetPaAcc := procedure
  begin
  	pEle := nil;
    hr := iVW.GetParentElement(cEle, pEle);
    //hr := cEle.GetCachedParent(pEle);

  end;

  GetFCld := procedure
  begin
  	fcldEle := nil;
    Scope := TreeScope_Children;
    hr := tEle.FindFirst(Scope, uiCondition, fcldEle);
  end;

  NodeMove := procedure
  begin
  	SSelNode.MoveTo(cNode, naAddChild);
    SSelNode := cNode;
    if SSelNode.Level = 0 then
    	SelPaNode := SSelNode.Item[iTarg];
  end;

  if Terminated then
    Exit;

  cEle := UiEle;
  SyncEle := cEle;
  pNode := SelNode;
  cNode := pNode;

		Synchronize(SCld);
		if hr_SCld = 0 then
		begin
			for i := 0 to integer(iLen) - 1 do
			begin
				if Terminated then
					break;
				sleep(1);
        itarg := i;
        Synchronize(GetCld);
        Synchronize(GetFCld);
        if Assigned(fcldEle) then
        begin
        	Synchronize(AddDummyNode);
        end
        else
        	rNode := AddTreeItem(SyncEle, pNode);
        Sleep(1);
			end;
		end;

	if bRecursive and (not Terminated) then
  begin

    SSelNode := Selnode;
    compEle := UiEle;
    Synchronize(GetPaAcc);

    while ((hr = S_OK) and Assigned(pEle)) do
    begin

      if Terminated then
      	break;
    	SyncEle := pEle;
      cEle := pEle;
      cNode := nil;
      rNode := AddTreeItem(SyncEle, cNode);
      bFSame := False;
      cNode := rNode;

			Synchronize(SCld);
			if hr_SCld = 0 then
			begin
				for i := 0 to integer(iLen) - 1 do
				begin
					if Terminated then
						break;
					sleep(1);
          itarg := i;
          Synchronize(GetCld);
        	Synchronize(GetFCld);

          if (not bFSame) and (IsSameUIElement(compEle, tEle)) then
          begin
          	itarg := i;
          	Synchronize(NodeMove );
            bFSame := True;
            compEle := pEle;
          end
          else
          begin
          	if Assigned(fcldEle) then
						begin
							Synchronize(AddDummyNode);
						end
						else
						begin
							rNode := AddTreeItem(tEle, cNode);
						end;
          end;
          sleep(1);
				end;   //for end
			end;


      if (IsSameUIElement(pEle, UIpEle)) then  break;
      Synchronize(GetPaAcc);
      sleep(1);
		end; //while end
  end;
end;


procedure TreeThread4UIA.RecurACC(ParentNode: TTreeNode; ParentEle: IUIAutomationElement; uiCondition: IUIAutomationCondition);
var
	tEle, SyncEle, fcldEle: IUIAutomationElement;
	iImgIdx: integer;
	cNode, Recurnode: TTreeNode;
	ws: widestring;
	SCld, GetCld, GetFCld: TThreadProcedure;
	ov: OleVariant;
	hr, hr_SCld: HResult;
	PC: PChar;
	arElement: IUIAutomationElementArray;
	iLen, i, iTarg: integer;
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

  SCld := procedure
  begin
  	Scope := TreeScope_Children;
		hr_SCld := SyncEle.FindAll(Scope, uiCondition, arElement);
		arElement.Get_Length(iLen);

  end;

  GetCld := procedure
  begin
  	tEle := nil;
    hr := arElement.GetElement(itarg, tEle);
  end;


  GetFCld := procedure
  begin
  	fcldEle := nil;
    Scope := TreeScope_Children;
    hr := tEle.FindFirst(Scope, uiCondition, fcldEle);
  end;

  cNode := ParentNode;
	SyncEle := ParentEle;
  if Assigned(cNode) and (cNode.Level >= 300) and Assigned(frmMSAAV.LoopNode) then
  	exit;


  if hr = S_OK then
  begin
    Synchronize(SCld);
    for i := 0 to iLen - 1 do
    begin
    	if Terminated then
				break;
			sleep(1);
			iTarg := i;
			Synchronize(GetCld);
			Synchronize(GetFCld);

			if Assigned(fcldEle) then
			begin
				Recurnode := AddTreeItem(tEle, cNode);
				RecurACC(Recurnode, tEle, uiCondition);
			end
			else
			begin
				AddTreeItem(tEle, cNode);
			end;
		end;
	end;

end;

end.
