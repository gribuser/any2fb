unit any_2_fb2_TLB;

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

// PASTLWTR : $Revision:   1.130  $
// File generated on 30.01.2005 2:07:13 from Type Library described below.

// ************************************************************************  //
// Type Lib: H:\Work\Any2FB\src\any_2_fb2.tlb (1)
// LIBID: {AC9A272F-112E-4F5E-ACF7-D975FA9B01CA}
// LCID: 0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (H:\WINDOWS\system32\stdole2.tlb)
//   (2) v4.0 StdVCL, (H:\WINDOWS\system32\stdvcl40.dll)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}

interface

uses ActiveX, Classes, Graphics, StdVCL, Variants, Windows;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  any_2_fb2MajorVersion = 1;
  any_2_fb2MinorVersion = 0;

  LIBID_any_2_fb2: TGUID = '{AC9A272F-112E-4F5E-ACF7-D975FA9B01CA}';

  IID_IFBEImportPlugin: TGUID = '{8094BC55-99C0-4ADF-BD55-71E206DFD403}';
  CLASS_FBEImportPlugin: TGUID = '{D3DF1A61-D99C-4794-BFBA-027A18F58780}';
  IID_IAny2FB2: TGUID = '{0233407F-7AA1-4B94-B263-5BE1CBFB31D3}';
  CLASS_Any2FB2: TGUID = '{A24186BE-6084-4DF6-9F63-5FD9B4B2A199}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IFBEImportPlugin = interface;
  IAny2FB2 = interface;
  IAny2FB2Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  FBEImportPlugin = IFBEImportPlugin;
  Any2FB2 = IAny2FB2;


// *********************************************************************//
// Interface: IFBEImportPlugin
// Flags:     (256) OleAutomation
// GUID:      {8094BC55-99C0-4ADF-BD55-71E206DFD403}
// *********************************************************************//
  IFBEImportPlugin = interface(IUnknown)
    ['{8094BC55-99C0-4ADF-BD55-71E206DFD403}']
    function  import(hWnd: Integer; out filename: WideString; out document: IDispatch): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IAny2FB2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0233407F-7AA1-4B94-B263-5BE1CBFB31D3}
// *********************************************************************//
  IAny2FB2 = interface(IDispatch)
    ['{0233407F-7AA1-4B94-B263-5BE1CBFB31D3}']
    function  Convert(URL: OleVariant): IDispatch; safecall;
    function  Get_PreserveForm: WordBool; safecall;
    procedure Set_PreserveForm(Value: WordBool); safecall;
    function  Get_noConvertCharset: WordBool; safecall;
    procedure Set_noConvertCharset(Value: WordBool); safecall;
    function  Get_noEpigraphs: WordBool; safecall;
    procedure Set_noEpigraphs(Value: WordBool); safecall;
    function  Get_noEmptyLines: WordBool; safecall;
    procedure Set_noEmptyLines(Value: WordBool); safecall;
    function  Get_noDescription: WordBool; safecall;
    procedure Set_noDescription(Value: WordBool); safecall;
    function  Get_FixCount: Integer; safecall;
    procedure Set_FixCount(Value: Integer); safecall;
    function  Get_noQuotesConvertion: WordBool; safecall;
    procedure Set_noQuotesConvertion(Value: WordBool); safecall;
    function  Get_noFootNotes: WordBool; safecall;
    procedure Set_noFootNotes(Value: WordBool); safecall;
    function  Get_noItalic: WordBool; safecall;
    procedure Set_noItalic(Value: WordBool); safecall;
    function  Get_noRestoreBrokenParagraphs: WordBool; safecall;
    procedure Set_noRestoreBrokenParagraphs(Value: WordBool); safecall;
    function  Get_noPoems: WordBool; safecall;
    procedure Set_noPoems(Value: WordBool); safecall;
    function  ConvertInteractive(hWnd: Integer; needSave: WordBool): IDispatch; safecall;
    function  Get_noHeaders: WordBool; safecall;
    procedure Set_noHeaders(Value: WordBool); safecall;
    function  Get_ignoreLineIndent: WordBool; safecall;
    procedure Set_ignoreLineIndent(Value: WordBool); safecall;
    function  Get_noLongDashes: WordBool; safecall;
    procedure Set_noLongDashes(Value: WordBool); safecall;
    procedure Set_TextType(Param1: Integer); safecall;
    function  Get_noImages: WordBool; safecall;
    procedure Set_noImages(Value: WordBool); safecall;
    function  Get_noOffSiteImages: WordBool; safecall;
    procedure Set_noOffSiteImages(Value: WordBool); safecall;
    function  Get_leaveDinamicImages: WordBool; safecall;
    procedure Set_leaveDinamicImages(Value: WordBool); safecall;
    function  Get_noExternalLinks: WordBool; safecall;
    procedure Set_noExternalLinks(Value: WordBool); safecall;
    function  Get_FollowLinksDeep: Integer; safecall;
    procedure Set_FollowLinksDeep(Value: Integer); safecall;
    function  Get_FollowOffSiteLinks: WordBool; safecall;
    procedure Set_FollowOffSiteLinks(Value: WordBool); safecall;
    function  Get_reOnlyFollowLinks: OleVariant; safecall;
    procedure Set_reOnlyFollowLinks(Value: OleVariant); safecall;
    function  Get_reNeverFollowLinks: OleVariant; safecall;
    procedure Set_reNeverFollowLinks(Value: OleVariant); safecall;
    function  Get_reHeadersDetect: OleVariant; safecall;
    procedure Set_reHeadersDetect(Value: OleVariant); safecall;
    function  Get_reOnLoad: OleVariant; safecall;
    procedure Set_reOnLoad(Value: OleVariant); safecall;
    function  Get_reOnDone: OleVariant; safecall;
    procedure Set_reOnDone(Value: OleVariant); safecall;
    function  Get_LOG: OleVariant; safecall;
    property PreserveForm: WordBool read Get_PreserveForm write Set_PreserveForm;
    property noConvertCharset: WordBool read Get_noConvertCharset write Set_noConvertCharset;
    property noEpigraphs: WordBool read Get_noEpigraphs write Set_noEpigraphs;
    property noEmptyLines: WordBool read Get_noEmptyLines write Set_noEmptyLines;
    property noDescription: WordBool read Get_noDescription write Set_noDescription;
    property FixCount: Integer read Get_FixCount write Set_FixCount;
    property noQuotesConvertion: WordBool read Get_noQuotesConvertion write Set_noQuotesConvertion;
    property noFootNotes: WordBool read Get_noFootNotes write Set_noFootNotes;
    property noItalic: WordBool read Get_noItalic write Set_noItalic;
    property noRestoreBrokenParagraphs: WordBool read Get_noRestoreBrokenParagraphs write Set_noRestoreBrokenParagraphs;
    property noPoems: WordBool read Get_noPoems write Set_noPoems;
    property noHeaders: WordBool read Get_noHeaders write Set_noHeaders;
    property ignoreLineIndent: WordBool read Get_ignoreLineIndent write Set_ignoreLineIndent;
    property noLongDashes: WordBool read Get_noLongDashes write Set_noLongDashes;
    property TextType: Integer write Set_TextType;
    property noImages: WordBool read Get_noImages write Set_noImages;
    property noOffSiteImages: WordBool read Get_noOffSiteImages write Set_noOffSiteImages;
    property leaveDinamicImages: WordBool read Get_leaveDinamicImages write Set_leaveDinamicImages;
    property noExternalLinks: WordBool read Get_noExternalLinks write Set_noExternalLinks;
    property FollowLinksDeep: Integer read Get_FollowLinksDeep write Set_FollowLinksDeep;
    property FollowOffSiteLinks: WordBool read Get_FollowOffSiteLinks write Set_FollowOffSiteLinks;
    property reOnlyFollowLinks: OleVariant read Get_reOnlyFollowLinks write Set_reOnlyFollowLinks;
    property reNeverFollowLinks: OleVariant read Get_reNeverFollowLinks write Set_reNeverFollowLinks;
    property reHeadersDetect: OleVariant read Get_reHeadersDetect write Set_reHeadersDetect;
    property reOnLoad: OleVariant read Get_reOnLoad write Set_reOnLoad;
    property reOnDone: OleVariant read Get_reOnDone write Set_reOnDone;
    property LOG: OleVariant read Get_LOG;
  end;

// *********************************************************************//
// DispIntf:  IAny2FB2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0233407F-7AA1-4B94-B263-5BE1CBFB31D3}
// *********************************************************************//
  IAny2FB2Disp = dispinterface
    ['{0233407F-7AA1-4B94-B263-5BE1CBFB31D3}']
    function  Convert(URL: OleVariant): IDispatch; dispid 1;
    property PreserveForm: WordBool dispid 2;
    property noConvertCharset: WordBool dispid 3;
    property noEpigraphs: WordBool dispid 4;
    property noEmptyLines: WordBool dispid 5;
    property noDescription: WordBool dispid 6;
    property FixCount: Integer dispid 7;
    property noQuotesConvertion: WordBool dispid 8;
    property noFootNotes: WordBool dispid 9;
    property noItalic: WordBool dispid 10;
    property noRestoreBrokenParagraphs: WordBool dispid 11;
    property noPoems: WordBool dispid 12;
    function  ConvertInteractive(hWnd: Integer; needSave: WordBool): IDispatch; dispid 13;
    property noHeaders: WordBool dispid 14;
    property ignoreLineIndent: WordBool dispid 15;
    property noLongDashes: WordBool dispid 16;
    property TextType: Integer writeonly dispid 17;
    property noImages: WordBool dispid 18;
    property noOffSiteImages: WordBool dispid 19;
    property leaveDinamicImages: WordBool dispid 20;
    property noExternalLinks: WordBool dispid 21;
    property FollowLinksDeep: Integer dispid 22;
    property FollowOffSiteLinks: WordBool dispid 23;
    property reOnlyFollowLinks: OleVariant dispid 24;
    property reNeverFollowLinks: OleVariant dispid 25;
    property reHeadersDetect: OleVariant dispid 26;
    property reOnLoad: OleVariant dispid 27;
    property reOnDone: OleVariant dispid 29;
    property LOG: OleVariant readonly dispid 28;
  end;

// *********************************************************************//
// The Class CoFBEImportPlugin provides a Create and CreateRemote method to          
// create instances of the default interface IFBEImportPlugin exposed by              
// the CoClass FBEImportPlugin. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFBEImportPlugin = class
    class function Create: IFBEImportPlugin;
    class function CreateRemote(const MachineName: string): IFBEImportPlugin;
  end;

// *********************************************************************//
// The Class CoAny2FB2 provides a Create and CreateRemote method to          
// create instances of the default interface IAny2FB2 exposed by              
// the CoClass Any2FB2. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoAny2FB2 = class
    class function Create: IAny2FB2;
    class function CreateRemote(const MachineName: string): IAny2FB2;
  end;

implementation

uses ComObj;

class function CoFBEImportPlugin.Create: IFBEImportPlugin;
begin
  Result := CreateComObject(CLASS_FBEImportPlugin) as IFBEImportPlugin;
end;

class function CoFBEImportPlugin.CreateRemote(const MachineName: string): IFBEImportPlugin;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FBEImportPlugin) as IFBEImportPlugin;
end;

class function CoAny2FB2.Create: IAny2FB2;
begin
  Result := CreateComObject(CLASS_Any2FB2) as IAny2FB2;
end;

class function CoAny2FB2.CreateRemote(const MachineName: string): IAny2FB2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Any2FB2) as IAny2FB2;
end;

end.
