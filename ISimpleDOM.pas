unit ISimpleDOM;

interface
uses Windows, ActiveX;
const
	NODETYPE_ELEMENT	= 1;
	NODETYPE_ATTRIBUTE	= 2;
	NODETYPE_TEXT	= 3;
	NODETYPE_CDATA_SECTION	= 4;
	NODETYPE_ENTITY_REFERENCE	= 5;
	NODETYPE_ENTITY	=  ( 6 );
	NODETYPE_PROCESSING_INSTRUCTION	= 7;
	NODETYPE_COMMENT	= 8;
	NODETYPE_DOCUMENT	= 9;
	NODETYPE_DOCUMENT_TYPE	= 10;
	NODETYPE_DOCUMENT_FRAGMENT	= 11;
	NODETYPE_NOTATION	= 12;
	IID_ISIMPLEDOMNODE: TGUID = '{1814ceeb-49e2-407f-af99-fa755a7d2607}';
	IID_ISIMPLEDOMDOCUMENT: TGUID = '{0D68D6D0-D93D-4d08-A30D-F00DD1F45B24}';
	IID_ISIMPLEDOMTEXT: TGUID = '{4e747be5-2052-4265-8af0-8ecad7aad1c0}';
        
        
type
  ISimpleDOMNode = interface ( IUnknown )
	['{1814ceeb-49e2-407f-af99-fa755a7d2607}']
	function get_nodeInfo(out nodeName: TBSTR; out nameSpaceID: Smallint; out nodeValue: TBSTR; out numChildren: PUINT; out uniqueID: PUINT;out nodeType: WORD): HRESULT; stdcall;
	function get_attributes(maxAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR; out numAttribs: WORD): HRESULT; stdcall;
	function get_attributesForNames(numAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR): HRESULT; stdcall;
	function get_computedStyle(maxStyleProperties: WORD; useAlternateView: boolean; out styleProperties: TBSTR; out styleValues: TBSTR; out numStyleProperties: WORD): HRESULT; stdcall;
	function get_computedStyleForProperties(numStyleProperties: WORD; useAlternateView: boolean; out styleProperties: TBSTR; out styleValues: TBSTR): HRESULT; stdcall;
	function scrollTo(placeTopLeft: boolean): HRESULT; stdcall;
	function get_parentNode(out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_firstChild(out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_lastChild(out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_previousSibling(out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_nextSibling(out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_childAt(childIndex: UINT; out node: ISimpleDOMNode): HRESULT; stdcall;
	function get_innerHTML(out innerHTML: TBSTR): HRESULT; stdcall;
	function get_localInterface(out localInterface: Pointer): HRESULT; stdcall;
	function get_language(out language: TBSTR): HRESULT; stdcall;
  end;
  ISimpleDOMDocument = interface ( IUnknown )
	['{0D68D6D0-D93D-4d08-A30D-F00DD1F45B24}']
	function get_URL(out url: TBSTR): HRESULT; stdcall;
	function get_title(out title: TBSTR): HRESULT; stdcall;
	function get_mimeType(out mimeType: TBSTR): HRESULT; stdcall;
	function get_docType(out docType: TBSTR): HRESULT; stdcall;
	function get_nameSpaceURIForID(nameSpaceID: Smallint; out nameSpaceURI: TBSTR): HRESULT; stdcall;
	function put_alternateViewMediaTypes(var commaSeparatedMediaTypes: TBSTR): HRESULT; stdcall;
  end;
  ISimpleDOMText = interface ( IUnknown )
	['{4e747be5-2052-4265-8af0-8ecad7aad1c0}']
	function get_domText(out domText: TBSTR): HRESULT; stdcall;
	function get_clippedSuTBSTRingBounds(startIndex: UINT; endIndex: UINT; out x: Integer; out y: Integer; out width: Integer; out height: Integer): HRESULT; stdcall;
	function get_unclippedSuTBSTRingBounds(startIndex: UINT; endIndex: UINT; out x: Integer; out y: Integer; out width: Integer; out height: Integer): HRESULT; stdcall;
	function scrollToSuTBSTRing(startIndex: UINT; endIndex: UINT): HRESULT; stdcall;
	function get_fontFamily(out fontFamily: TBSTR): HRESULT; stdcall;
  end;
  
implementation

end.