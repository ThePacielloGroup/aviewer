unit IAccessible2Lib_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 8291 $
// File generated on 2010/03/09 19:29:20 from Type Library described below.

// ************************************************************************  //
// Type Lib: M:\Gekco\ia2_api_all.tlb (1)
// LIBID: {21955A1E-5C3A-4E52-8CCB-CE587E75133E}
// LCID: 0
// Helpfile: 
// HelpString: IAccessible2 Type Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
// Errors:
//   Hint: Member 'end' of 'IA2TextSegment' changed to 'end_'
//   Hint: Symbol 'type' renamed to 'type_'
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants, oleacc;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  IAccessible2LibMajorVersion = 1;
  IAccessible2LibMinorVersion = 0;

  LIBID_IAccessible2Lib: TGUID = '{21955A1E-5C3A-4E52-8CCB-CE587E75133E}';

  //IID_IAccessible: TGUID = '{618736E0-3C3D-11CF-810C-00AA00389B71}';
  IID_IAccessible2: TGUID = '{E89F726E-C4F4-4C19-BB19-B647D7FA8478}';
  IID_IAccessibleRelation: TGUID = '{7CDF86EE-C3DA-496A-BDA4-281B336E1FDC}';
  IID_IAccessibleAction: TGUID = '{B70D9F59-3B5A-4DBA-AB9E-22012F607DF5}';
  IID_IAccessibleApplication: TGUID = '{D49DED83-5B25-43F4-9B95-93B44595979E}';
  IID_IAccessibleComponent: TGUID = '{1546D4B0-4C98-4BDA-89AE-9A64748BDDE4}';
  IID_IAccessibleEditableText: TGUID = '{A59AA09A-7011-4B65-939D-32B1FB5547E3}';
  IID_IAccessibleHyperlink: TGUID = '{01C20F2B-3DD2-400F-949F-AD00BDAB1D41}';
  IID_IAccessibleText: TGUID = '{24FD2FFB-3AAD-4A08-8335-A3AD89C0FB4B}';
  IID_IAccessibleHypertext: TGUID = '{6B4F8BBF-F1F2-418A-B35E-A195BC4103B9}';
  IID_IAccessibleImage: TGUID = '{FE5ABB3D-615E-4F7B-909F-5F0EDA9E8DDE}';
  IID_IAccessibleTable: TGUID = '{35AD8070-C20C-4FB4-B094-F4F7275DD469}';
  IID_IAccessibleValue: TGUID = '{35855B5B-C566-4FD0-A7B1-E65465600394}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum IA2ScrollType
type
  IA2ScrollType = TOleEnum;
const
  IA2_SCROLL_TYPE_TOP_LEFT = $00000000;
  IA2_SCROLL_TYPE_BOTTOM_RIGHT = $00000001;
  IA2_SCROLL_TYPE_TOP_EDGE = $00000002;
  IA2_SCROLL_TYPE_BOTTOM_EDGE = $00000003;
  IA2_SCROLL_TYPE_LEFT_EDGE = $00000004;
  IA2_SCROLL_TYPE_RIGHT_EDGE = $00000005;
  IA2_SCROLL_TYPE_ANYWHERE = $00000006;

// Constants for enum IA2CoordinateType
type
  IA2CoordinateType = TOleEnum;
const
  IA2_COORDTYPE_SCREEN_RELATIVE = $00000000;
  IA2_COORDTYPE_PARENT_RELATIVE = $00000001;

// Constants for enum IA2TextBoundaryType
type
  IA2TextBoundaryType = TOleEnum;
const
  IA2_TEXT_BOUNDARY_CHAR = $00000000;
  IA2_TEXT_BOUNDARY_WORD = $00000001;
  IA2_TEXT_BOUNDARY_SENTENCE = $00000002;
  IA2_TEXT_BOUNDARY_PARAGRAPH = $00000003;
  IA2_TEXT_BOUNDARY_LINE = $00000004;
  IA2_TEXT_BOUNDARY_ALL = $00000005;

// Constants for enum IA2TableModelChangeType
type
  IA2TableModelChangeType = TOleEnum;
const
  IA2_TABLE_MODEL_CHANGE_INSERT = $00000000;
  IA2_TABLE_MODEL_CHANGE_DELETE = $00000001;
  IA2_TABLE_MODEL_CHANGE_UPDATE = $00000002;

// Constants for enum IA2EventID
type
  IA2EventID = TOleEnum;
const
  IA2_EVENT_ACTION_CHANGED = $00000101;
  IA2_EVENT_ACTIVE_DECENDENT_CHANGED = $00000102;
  IA2_EVENT_ACTIVE_DESCENDANT_CHANGED = $00000102;
  IA2_EVENT_DOCUMENT_ATTRIBUTE_CHANGED = $00000103;
  IA2_EVENT_DOCUMENT_CONTENT_CHANGED = $00000104;
  IA2_EVENT_DOCUMENT_LOAD_COMPLETE = $00000105;
  IA2_EVENT_DOCUMENT_LOAD_STOPPED = $00000106;
  IA2_EVENT_DOCUMENT_RELOAD = $00000107;
  IA2_EVENT_HYPERLINK_END_INDEX_CHANGED = $00000108;
  IA2_EVENT_HYPERLINK_NUMBER_OF_ANCHORS_CHANGED = $00000109;
  IA2_EVENT_HYPERLINK_SELECTED_LINK_CHANGED = $0000010A;
  IA2_EVENT_HYPERTEXT_LINK_ACTIVATED = $0000010B;
  IA2_EVENT_HYPERTEXT_LINK_SELECTED = $0000010C;
  IA2_EVENT_HYPERLINK_START_INDEX_CHANGED = $0000010D;
  IA2_EVENT_HYPERTEXT_CHANGED = $0000010E;
  IA2_EVENT_HYPERTEXT_NLINKS_CHANGED = $0000010F;
  IA2_EVENT_OBJECT_ATTRIBUTE_CHANGED = $00000110;
  IA2_EVENT_PAGE_CHANGED = $00000111;
  IA2_EVENT_SECTION_CHANGED = $00000112;
  IA2_EVENT_TABLE_CAPTION_CHANGED = $00000113;
  IA2_EVENT_TABLE_COLUMN_DESCRIPTION_CHANGED = $00000114;
  IA2_EVENT_TABLE_COLUMN_HEADER_CHANGED = $00000115;
  IA2_EVENT_TABLE_MODEL_CHANGED = $00000116;
  IA2_EVENT_TABLE_ROW_DESCRIPTION_CHANGED = $00000117;
  IA2_EVENT_TABLE_ROW_HEADER_CHANGED = $00000118;
  IA2_EVENT_TABLE_SUMMARY_CHANGED = $00000119;
  IA2_EVENT_TEXT_ATTRIBUTE_CHANGED = $0000011A;
  IA2_EVENT_TEXT_CARET_MOVED = $0000011B;
  IA2_EVENT_TEXT_CHANGED = $0000011C;
  IA2_EVENT_TEXT_COLUMN_CHANGED = $0000011D;
  IA2_EVENT_TEXT_INSERTED = $0000011E;
  IA2_EVENT_TEXT_REMOVED = $0000011F;
  IA2_EVENT_TEXT_UPDATED = $00000120;
  IA2_EVENT_TEXT_SELECTION_CHANGED = $00000121;
  IA2_EVENT_VISIBLE_DATA_CHANGED = $00000122;

// Constants for enum IA2Role
type
  IA2Role = TOleEnum;
const
  IA2_ROLE_UNKNOWN = $00000000;
  IA2_ROLE_CANVAS = $00000401;
  IA2_ROLE_CAPTION = $00000402;
  IA2_ROLE_CHECK_MENU_ITEM = $00000403;
  IA2_ROLE_COLOR_CHOOSER = $00000404;
  IA2_ROLE_DATE_EDITOR = $00000405;
  IA2_ROLE_DESKTOP_ICON = $00000406;
  IA2_ROLE_DESKTOP_PANE = $00000407;
  IA2_ROLE_DIRECTORY_PANE = $00000408;
  IA2_ROLE_EDITBAR = $00000409;
  IA2_ROLE_EMBEDDED_OBJECT = $0000040A;
  IA2_ROLE_ENDNOTE = $0000040B;
  IA2_ROLE_FILE_CHOOSER = $0000040C;
  IA2_ROLE_FONT_CHOOSER = $0000040D;
  IA2_ROLE_FOOTER = $0000040E;
  IA2_ROLE_FOOTNOTE = $0000040F;
  IA2_ROLE_FORM = $00000410;
  IA2_ROLE_FRAME = $00000411;
  IA2_ROLE_GLASS_PANE = $00000412;
  IA2_ROLE_HEADER = $00000413;
  IA2_ROLE_HEADING = $00000414;
  IA2_ROLE_ICON = $00000415;
  IA2_ROLE_IMAGE_MAP = $00000416;
  IA2_ROLE_INPUT_METHOD_WINDOW = $00000417;
  IA2_ROLE_INTERNAL_FRAME = $00000418;
  IA2_ROLE_LABEL = $00000419;
  IA2_ROLE_LAYERED_PANE = $0000041A;
  IA2_ROLE_NOTE = $0000041B;
  IA2_ROLE_OPTION_PANE = $0000041C;
  IA2_ROLE_PAGE = $0000041D;
  IA2_ROLE_PARAGRAPH = $0000041E;
  IA2_ROLE_RADIO_MENU_ITEM = $0000041F;
  IA2_ROLE_REDUNDANT_OBJECT = $00000420;
  IA2_ROLE_ROOT_PANE = $00000421;
  IA2_ROLE_RULER = $00000422;
  IA2_ROLE_SCROLL_PANE = $00000423;
  IA2_ROLE_SECTION = $00000424;
  IA2_ROLE_SHAPE = $00000425;
  IA2_ROLE_SPLIT_PANE = $00000426;
  IA2_ROLE_TEAR_OFF_MENU = $00000427;
  IA2_ROLE_TERMINAL = $00000428;
  IA2_ROLE_TEXT_FRAME = $00000429;
  IA2_ROLE_TOGGLE_BUTTON = $0000042A;
  IA2_ROLE_VIEW_PORT = $0000042B;

// Constants for enum IA2States
type
  IA2States = TOleEnum;
const
  IA2_STATE_ACTIVE = $00000001;
  IA2_STATE_ARMED = $00000002;
  IA2_STATE_DEFUNCT = $00000004;
  IA2_STATE_EDITABLE = $00000008;
  IA2_STATE_HORIZONTAL = $00000010;
  IA2_STATE_ICONIFIED = $00000020;
  IA2_STATE_INVALID_ENTRY = $00000040;
  IA2_STATE_MANAGES_DESCENDANTS = $00000080;
  IA2_STATE_MODAL = $00000100;
  IA2_STATE_MULTI_LINE = $00000200;
  IA2_STATE_OPAQUE = $00000400;
  IA2_STATE_REQUIRED = $00000800;
  IA2_STATE_SELECTABLE_TEXT = $00001000;
  IA2_STATE_SINGLE_LINE = $00002000;
  IA2_STATE_STALE = $00004000;
  IA2_STATE_SUPPORTS_AUTOCOMPLETION = $00008000;
  IA2_STATE_TRANSIENT = $00010000;
  IA2_STATE_VERTICAL = $00020000;

// Constants for enum IA2TextSpecialOffsets
type
  IA2TextSpecialOffsets = TOleEnum;
const
  IA2_TEXT_OFFSET_LENGTH = $FFFFFFFF;
  IA2_TEXT_OFFSET_CARET = $FFFFFFFE;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  //IAccessible = interface;
  //IAccessibleDisp = dispinterface;
  IAccessible2 = interface;
  IAccessibleRelation = interface;
  IAccessibleAction = interface;
  IAccessibleApplication = interface;
  IAccessibleComponent = interface;
  IAccessibleEditableText = interface;
  IAccessibleHyperlink = interface;
  IAccessibleText = interface;
  IAccessibleHypertext = interface;
  IAccessibleImage = interface;
  IAccessibleTable = interface;
  IAccessibleValue = interface;

// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  wireHWND = ^_RemotableHandle; 
  PWideString1 = ^WideString; {*}
  PInteger1 = ^Integer; {*}


  __MIDL_IWinTypes_0009 = record
    case Integer of
      0: (hInproc: Integer);
      1: (hRemote: Integer);
  end;

  _RemotableHandle = packed record
    fContext: Integer;
    u: __MIDL_IWinTypes_0009;
  end;

  IA2Locale = packed record
    language: WideString;
    country: WideString;
    variant: WideString;
  end;

  IA2TextSegment = packed record
    text: WideString;
    start: Integer;
    end_: Integer;
  end;

  IA2TableModelChange = packed record
    type_: IA2TableModelChangeType;
    firstRow: Integer;
    lastRow: Integer;
    firstColumn: Integer;
    lastColumn: Integer;
  end;


// *********************************************************************//
// Interface: IAccessible
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {618736E0-3C3D-11CF-810C-00AA00389B71}
// *********************************************************************//
(*  IAccessible = interface(IDispatch)
    ['{618736E0-3C3D-11CF-810C-00AA00389B71}']
    function Get_accParent: IDispatch; safecall;
    function Get_accChildCount: Integer; safecall;
    function Get_accChild(varChild: OleVariant): IDispatch; safecall;
    function Get_accName(varChild: OleVariant): WideString; safecall;
    function Get_accValue(varChild: OleVariant): WideString; safecall;
    function Get_accDescription(varChild: OleVariant): WideString; safecall;
    function Get_accRole(varChild: OleVariant): OleVariant; safecall;
    function Get_accState(varChild: OleVariant): OleVariant; safecall;
    function Get_accHelp(varChild: OleVariant): WideString; safecall;
    function Get_accHelpTopic(out pszHelpFile: WideString; varChild: OleVariant): Integer; safecall;
    function Get_accKeyboardShortcut(varChild: OleVariant): WideString; safecall;
    function Get_accFocus: OleVariant; safecall;
    function Get_accSelection: OleVariant; safecall;
    function Get_accDefaultAction(varChild: OleVariant): WideString; safecall;
    procedure accSelect(flagsSelect: Integer; varChild: OleVariant); safecall;
    procedure accLocation(out pxLeft: Integer; out pyTop: Integer; out pcxWidth: Integer; 
                          out pcyHeight: Integer; varChild: OleVariant); safecall;
    function accNavigate(navDir: Integer; varStart: OleVariant): OleVariant; safecall;
    function accHitTest(xLeft: Integer; yTop: Integer): OleVariant; safecall;
    procedure accDoDefaultAction(varChild: OleVariant); safecall;
    procedure Set_accName(varChild: OleVariant; const pszName: WideString); safecall;
    procedure Set_accValue(varChild: OleVariant; const pszValue: WideString); safecall;
    property accParent: IDispatch read Get_accParent;
    property accChildCount: Integer read Get_accChildCount;
    property accChild[varChild: OleVariant]: IDispatch read Get_accChild;
    property accName[varChild: OleVariant]: WideString read Get_accName write Set_accName;
    property accValue[varChild: OleVariant]: WideString read Get_accValue write Set_accValue;
    property accDescription[varChild: OleVariant]: WideString read Get_accDescription;
    property accRole[varChild: OleVariant]: OleVariant read Get_accRole;
    property accState[varChild: OleVariant]: OleVariant read Get_accState;
    property accHelp[varChild: OleVariant]: WideString read Get_accHelp;
    property accHelpTopic[out pszHelpFile: WideString; varChild: OleVariant]: Integer read Get_accHelpTopic;
    property accKeyboardShortcut[varChild: OleVariant]: WideString read Get_accKeyboardShortcut;
    property accFocus: OleVariant read Get_accFocus;
    property accSelection: OleVariant read Get_accSelection;
    property accDefaultAction[varChild: OleVariant]: WideString read Get_accDefaultAction;
  end;
*)

// *********************************************************************//
// DispIntf:  IAccessibleDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {618736E0-3C3D-11CF-810C-00AA00389B71}
// *********************************************************************//
(*  IAccessibleDisp = dispinterface
    ['{618736E0-3C3D-11CF-810C-00AA00389B71}']
    property accParent: IDispatch readonly dispid -5000;
    property accChildCount: Integer readonly dispid -5001;
    property accChild[varChild: OleVariant]: IDispatch readonly dispid -5002;
    property accName[varChild: OleVariant]: WideString dispid -5003;
    property accValue[varChild: OleVariant]: WideString dispid -5004;
    property accDescription[varChild: OleVariant]: WideString readonly dispid -5005;
    property accRole[varChild: OleVariant]: OleVariant readonly dispid -5006;
    property accState[varChild: OleVariant]: OleVariant readonly dispid -5007;
    property accHelp[varChild: OleVariant]: WideString readonly dispid -5008;
    property accHelpTopic[out pszHelpFile: WideString; varChild: OleVariant]: Integer readonly dispid -5009;
    property accKeyboardShortcut[varChild: OleVariant]: WideString readonly dispid -5010;
    property accFocus: OleVariant readonly dispid -5011;
    property accSelection: OleVariant readonly dispid -5012;
    property accDefaultAction[varChild: OleVariant]: WideString readonly dispid -5013;
    procedure accSelect(flagsSelect: Integer; varChild: OleVariant); dispid -5014;
    procedure accLocation(out pxLeft: Integer; out pyTop: Integer; out pcxWidth: Integer; 
                          out pcyHeight: Integer; varChild: OleVariant); dispid -5015;
    function accNavigate(navDir: Integer; varStart: OleVariant): OleVariant; dispid -5016;
    function accHitTest(xLeft: Integer; yTop: Integer): OleVariant; dispid -5017;
    procedure accDoDefaultAction(varChild: OleVariant); dispid -5018;
  end;
*)
// *********************************************************************//
// Interface: IAccessible2
// Flags:     (4096) Dispatchable
// GUID:      {E89F726E-C4F4-4C19-BB19-B647D7FA8478}
// *********************************************************************//
  IAccessible2 = interface(IAccessible)
    ['{E89F726E-C4F4-4C19-BB19-B647D7FA8478}']
    function Get_nRelations(out nRelations: Integer): HResult; stdcall;
    function Get_relation(relationIndex: Integer; out relation: IAccessibleRelation): HResult; stdcall;
    function Get_relations(maxRelations: Integer; out relations: IAccessibleRelation; 
                           out nRelations: Integer): HResult; stdcall;
    function role(out role: Integer): HResult; stdcall;
    function scrollTo(scrollType: IA2ScrollType): HResult; stdcall;
    function scrollToPoint(coordinateType: IA2CoordinateType; x: Integer; y: Integer): HResult; stdcall;
    function Get_groupPosition(out groupLevel: Integer; out similarItemsInGroup: Integer; 
                               out positionInGroup: Integer): HResult; stdcall;
    function Get_states(out states: Integer): HResult; stdcall;
    function Get_extendedRole(out extendedRole: WideString): HResult; stdcall;
    function Get_localizedExtendedRole(out localizedExtendedRole: WideString): HResult; stdcall;
    function Get_nExtendedStates(out nExtendedStates: Integer): HResult; stdcall;
    function Get_extendedStates(maxExtendedStates: Integer; out extendedStates: PWideString1; 
                                out nExtendedStates: Integer): HResult; stdcall;
    function Get_localizedExtendedStates(maxLocalizedExtendedStates: Integer; 
                                         out localizedExtendedStates: PWideString1; 
                                         out nLocalizedExtendedStates: Integer): HResult; stdcall;
    function Get_uniqueID(out uniqueID: Integer): HResult; stdcall;
    function Get_windowHandle(out windowHandle: wireHWND): HResult; stdcall;
    function Get_indexInParent(out indexInParent: Integer): HResult; stdcall;
    function Get_locale(out locale: IA2Locale): HResult; stdcall;
    function Get_attributes(out attributes: WideString): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleRelation
// Flags:     (0)
// GUID:      {7CDF86EE-C3DA-496A-BDA4-281B336E1FDC}
// *********************************************************************//
  IAccessibleRelation = interface(IUnknown)
    ['{7CDF86EE-C3DA-496A-BDA4-281B336E1FDC}']
    function Get_relationType(out relationType: WideString): HResult; stdcall;
    function Get_localizedRelationType(out localizedRelationType: WideString): HResult; stdcall;
    function Get_nTargets(out nTargets: Integer): HResult; stdcall;
    function Get_target(targetIndex: Integer; out target: IUnknown): HResult; stdcall;
    function Get_targets(maxTargets: Integer; out targets: IUnknown; out nTargets: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleAction
// Flags:     (0)
// GUID:      {B70D9F59-3B5A-4DBA-AB9E-22012F607DF5}
// *********************************************************************//
  IAccessibleAction = interface(IUnknown)
    ['{B70D9F59-3B5A-4DBA-AB9E-22012F607DF5}']
    function nActions(out nActions: Integer): HResult; stdcall;
    function doAction(actionIndex: Integer): HResult; stdcall;
    function Get_description(actionIndex: Integer; out description: WideString): HResult; stdcall;
    function Get_keyBinding(actionIndex: Integer; nMaxBindings: Integer; 
                            out keyBindings: PWideString1; out nBindings: Integer): HResult; stdcall;
    function Get_name(actionIndex: Integer; out name: WideString): HResult; stdcall;
    function Get_localizedName(actionIndex: Integer; out localizedName: WideString): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleApplication
// Flags:     (0)
// GUID:      {D49DED83-5B25-43F4-9B95-93B44595979E}
// *********************************************************************//
  IAccessibleApplication = interface(IUnknown)
    ['{D49DED83-5B25-43F4-9B95-93B44595979E}']
    function Get_appName(out name: WideString): HResult; stdcall;
    function Get_appVersion(out version: WideString): HResult; stdcall;
    function Get_toolkitName(out name: WideString): HResult; stdcall;
    function Get_toolkitVersion(out version: WideString): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleComponent
// Flags:     (0)
// GUID:      {1546D4B0-4C98-4BDA-89AE-9A64748BDDE4}
// *********************************************************************//
  IAccessibleComponent = interface(IUnknown)
    ['{1546D4B0-4C98-4BDA-89AE-9A64748BDDE4}']
    function Get_locationInParent(out x: Integer; out y: Integer): HResult; stdcall;
    function Get_foreground(out foreground: Integer): HResult; stdcall;
    function Get_background(out background: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleEditableText
// Flags:     (0)
// GUID:      {A59AA09A-7011-4B65-939D-32B1FB5547E3}
// *********************************************************************//
  IAccessibleEditableText = interface(IUnknown)
    ['{A59AA09A-7011-4B65-939D-32B1FB5547E3}']
    function copyText(startOffset: Integer; endOffset: Integer): HResult; stdcall;
    function deleteText(startOffset: Integer; endOffset: Integer): HResult; stdcall;
    function insertText(offset: Integer; var text: WideString): HResult; stdcall;
    function cutText(startOffset: Integer; endOffset: Integer): HResult; stdcall;
    function pasteText(offset: Integer): HResult; stdcall;
    function replaceText(startOffset: Integer; endOffset: Integer; var text: WideString): HResult; stdcall;
    function setAttributes(startOffset: Integer; endOffset: Integer; var attributes: WideString): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleHyperlink
// Flags:     (0)
// GUID:      {01C20F2B-3DD2-400F-949F-AD00BDAB1D41}
// *********************************************************************//
  IAccessibleHyperlink = interface(IAccessibleAction)
    ['{01C20F2B-3DD2-400F-949F-AD00BDAB1D41}']
    function Get_anchor(index: Integer; out anchor: OleVariant): HResult; stdcall;
    function Get_anchorTarget(index: Integer; out anchorTarget: OleVariant): HResult; stdcall;
    function Get_startIndex(out index: Integer): HResult; stdcall;
    function Get_endIndex(out index: Integer): HResult; stdcall;
    function Get_valid(out valid: Shortint): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleText
// Flags:     (0)
// GUID:      {24FD2FFB-3AAD-4A08-8335-A3AD89C0FB4B}
// *********************************************************************//
  IAccessibleText = interface(IUnknown)
    ['{24FD2FFB-3AAD-4A08-8335-A3AD89C0FB4B}']
    function addSelection(startOffset: Integer; endOffset: Integer): HResult; stdcall;
    function Get_attributes(offset: Integer; out startOffset: Integer; out endOffset: Integer; 
                            out textAttributes: WideString): HResult; stdcall;
    function Get_caretOffset(out offset: Integer): HResult; stdcall;
    function Get_characterExtents(offset: Integer; coordType: IA2CoordinateType; out x: Integer; 
                                  out y: Integer; out width: Integer; out height: Integer): HResult; stdcall;
    function Get_nSelections(out nSelections: Integer): HResult; stdcall;
    function Get_offsetAtPoint(x: Integer; y: Integer; coordType: IA2CoordinateType; 
                               out offset: Integer): HResult; stdcall;
    function Get_selection(selectionIndex: Integer; out startOffset: Integer; out endOffset: Integer): HResult; stdcall;
    function Get_text(startOffset: Integer; endOffset: Integer; out text: WideString): HResult; stdcall;
    function Get_textBeforeOffset(offset: Integer; boundaryType: IA2TextBoundaryType; 
                                  out startOffset: Integer; out endOffset: Integer; 
                                  out text: WideString): HResult; stdcall;
    function Get_textAfterOffset(offset: Integer; boundaryType: IA2TextBoundaryType; 
                                 out startOffset: Integer; out endOffset: Integer; 
                                 out text: WideString): HResult; stdcall;
    function Get_textAtOffset(offset: Integer; boundaryType: IA2TextBoundaryType; 
                              out startOffset: Integer; out endOffset: Integer; out text: WideString): HResult; stdcall;
    function removeSelection(selectionIndex: Integer): HResult; stdcall;
    function setCaretOffset(offset: Integer): HResult; stdcall;
    function setSelection(selectionIndex: Integer; startOffset: Integer; endOffset: Integer): HResult; stdcall;
    function Get_nCharacters(out nCharacters: Integer): HResult; stdcall;
    function scrollSubstringTo(startIndex: Integer; endIndex: Integer; scrollType: IA2ScrollType): HResult; stdcall;
    function scrollSubstringToPoint(startIndex: Integer; endIndex: Integer; 
                                    coordinateType: IA2CoordinateType; x: Integer; y: Integer): HResult; stdcall;
    function Get_newText(out newText: IA2TextSegment): HResult; stdcall;
    function Get_oldText(out oldText: IA2TextSegment): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleHypertext
// Flags:     (0)
// GUID:      {6B4F8BBF-F1F2-418A-B35E-A195BC4103B9}
// *********************************************************************//
  IAccessibleHypertext = interface(IAccessibleText)
    ['{6B4F8BBF-F1F2-418A-B35E-A195BC4103B9}']
    function Get_nHyperlinks(out hyperlinkCount: Integer): HResult; stdcall;
    function Get_hyperlink(index: Integer; out hyperlink: IAccessibleHyperlink): HResult; stdcall;
    function Get_hyperlinkIndex(charIndex: Integer; out hyperlinkIndex: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleImage
// Flags:     (0)
// GUID:      {FE5ABB3D-615E-4F7B-909F-5F0EDA9E8DDE}
// *********************************************************************//
  IAccessibleImage = interface(IUnknown)
    ['{FE5ABB3D-615E-4F7B-909F-5F0EDA9E8DDE}']
    function Get_description(out description: WideString): HResult; stdcall;
    function Get_imagePosition(coordinateType: IA2CoordinateType; out x: Integer; out y: Integer): HResult; stdcall;
    function Get_imageSize(out height: Integer; out width: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleTable
// Flags:     (0)
// GUID:      {35AD8070-C20C-4FB4-B094-F4F7275DD469}
// *********************************************************************//
  IAccessibleTable = interface(IUnknown)
    ['{35AD8070-C20C-4FB4-B094-F4F7275DD469}']
    function Get_accessibleAt(row: Integer; column: Integer; out accessible: IUnknown): HResult; stdcall;
    function Get_caption(out accessible: IUnknown): HResult; stdcall;
    function Get_childIndex(rowIndex: Integer; columnIndex: Integer; out cellIndex: Integer): HResult; stdcall;
    function Get_columnDescription(column: Integer; out description: WideString): HResult; stdcall;
    function Get_columnExtentAt(row: Integer; column: Integer; out nColumnsSpanned: Integer): HResult; stdcall;
    function Get_columnHeader(out accessibleTable: IAccessibleTable; out startingRowIndex: Integer): HResult; stdcall;
    function Get_columnIndex(cellIndex: Integer; out columnIndex: Integer): HResult; stdcall;
    function Get_nColumns(out columnCount: Integer): HResult; stdcall;
    function Get_nRows(out rowCount: Integer): HResult; stdcall;
    function Get_nSelectedChildren(out cellCount: Integer): HResult; stdcall;
    function Get_nSelectedColumns(out columnCount: Integer): HResult; stdcall;
    function Get_nSelectedRows(out rowCount: Integer): HResult; stdcall;
    function Get_rowDescription(row: Integer; out description: WideString): HResult; stdcall;
    function Get_rowExtentAt(row: Integer; column: Integer; out nRowsSpanned: Integer): HResult; stdcall;
    function Get_rowHeader(out accessibleTable: IAccessibleTable; out startingColumnIndex: Integer): HResult; stdcall;
    function Get_rowIndex(cellIndex: Integer; out rowIndex: Integer): HResult; stdcall;
    function Get_selectedChildren(maxChildren: Integer; out children: PInteger1; 
                                  out nChildren: Integer): HResult; stdcall;
    function Get_selectedColumns(maxColumns: Integer; out columns: PInteger1; out nColumns: Integer): HResult; stdcall;
    function Get_selectedRows(maxRows: Integer; out rows: PInteger1; out nRows: Integer): HResult; stdcall;
    function Get_summary(out accessible: IUnknown): HResult; stdcall;
    function Get_isColumnSelected(column: Integer; out isSelected: Shortint): HResult; stdcall;
    function Get_isRowSelected(row: Integer; out isSelected: Shortint): HResult; stdcall;
    function Get_isSelected(row: Integer; column: Integer; out isSelected: Shortint): HResult; stdcall;
    function selectRow(row: Integer): HResult; stdcall;
    function selectColumn(column: Integer): HResult; stdcall;
    function unselectRow(row: Integer): HResult; stdcall;
    function unselectColumn(column: Integer): HResult; stdcall;
    function Get_rowColumnExtentsAtIndex(index: Integer; out row: Integer; out column: Integer; 
                                         out rowExtents: Integer; out columnExtents: Integer; 
                                         out isSelected: Shortint): HResult; stdcall;
    function Get_modelChange(out modelChange: IA2TableModelChange): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAccessibleValue
// Flags:     (0)
// GUID:      {35855B5B-C566-4FD0-A7B1-E65465600394}
// *********************************************************************//
  IAccessibleValue = interface(IUnknown)
    ['{35855B5B-C566-4FD0-A7B1-E65465600394}']
    function Get_currentValue(out currentValue: OleVariant): HResult; stdcall;
    function setCurrentValue(value: OleVariant): HResult; stdcall;
    function Get_maximumValue(out maximumValue: OleVariant): HResult; stdcall;
    function Get_minimumValue(out minimumValue: OleVariant): HResult; stdcall;
  end;

implementation

uses ComObj;

end.
