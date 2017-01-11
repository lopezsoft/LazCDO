Unit CDO_1_0_TLB;

//  Imported CDO on 23/12/2016 2:21:36 from C:\Windows\SysWOW64\cdosys.dll

{$mode delphi}{$H+}

interface

// Dependency: ADODB v6.0 (C:\Program Files (x86)\Common Files\System\ado\msado60.tlb)
//  Warning: renamed parameter 'Interface' in IBodyPart.GetInterface to 'Interface_'
//  Warning: renamed parameter 'Interface' in IBodyPart.GetInterface to 'Interface_'
//  Warning: renamed property 'To' in IMessage to 'To_'
//  Warning: renamed parameter 'Interface' in IMessage.GetInterface to 'Interface_'
//  Warning: renamed parameter 'Interface' in IMessage.GetInterface to 'Interface_'
//  Warning: renamed parameter 'Interface' in IConfiguration.GetInterface to 'Interface_'
//  Warning: renamed parameter 'Interface' in IConfiguration.GetInterface to 'Interface_'
Uses
  Windows,ActiveX,Classes,Variants,ADODB_6_0_TLB;
Const
  CDOMajorVersion = 1;
  CDOMinorVersion = 0;
  CDOLCID = 0;
  LIBID_CDO : TGUID = '{CD000000-8B95-11D1-82DB-00C04FB1625D}';

  IID_IBodyParts : TGUID = '{CD000023-8B95-11D1-82DB-00C04FB1625D}';
  IID_IBodyPart : TGUID = '{CD000021-8B95-11D1-82DB-00C04FB1625D}';
  IID_IDataSource : TGUID = '{CD000029-8B95-11D1-82DB-00C04FB1625D}';
  IID_IMessages : TGUID = '{CD000025-8B95-11D1-82DB-00C04FB1625D}';
  IID_IMessage : TGUID = '{CD000020-8B95-11D1-82DB-00C04FB1625D}';
  IID_IConfiguration : TGUID = '{CD000022-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_Message : TGUID = '{CD000001-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_Configuration : TGUID = '{CD000002-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_DropDirectory : TGUID = '{CD000004-8B95-11D1-82DB-00C04FB1625D}';
  IID_IDropDirectory : TGUID = '{CD000024-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_SMTPConnector : TGUID = '{CD000008-8B95-11D1-82DB-00C04FB1625D}';
  IID_ISMTPScriptConnector : TGUID = '{CD000030-8B95-11D1-82DB-00C04FB1625D}';
  IID_ISMTPOnArrival : TGUID = '{CD000026-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_NNTPEarlyConnector : TGUID = '{CD000011-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPEarlyScriptConnector : TGUID = '{CD000034-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPOnPostEarly : TGUID = '{CD000033-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_NNTPPostConnector : TGUID = '{CD000009-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPPostScriptConnector : TGUID = '{CD000031-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPOnPost : TGUID = '{CD000027-8B95-11D1-82DB-00C04FB1625D}';
  CLASS_NNTPFinalConnector : TGUID = '{CD000010-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPFinalScriptConnector : TGUID = '{CD000032-8B95-11D1-82DB-00C04FB1625D}';
  IID_INNTPOnPostFinal : TGUID = '{CD000028-8B95-11D1-82DB-00C04FB1625D}';

//Enums

Type
  CdoConfigSource =LongWord;
Const
  cdoDefaults = $FFFFFFFFFFFFFFFF;
  cdoIIS = $0000000000000001;
  cdoOutlookExpress = $0000000000000002;
Type
  CdoDSNOptions =LongWord;
Const
  cdoDSNDefault = $0000000000000000;
  cdoDSNNever = $0000000000000001;
  cdoDSNFailure = $0000000000000002;
  cdoDSNSuccess = $0000000000000004;
  cdoDSNDelay = $0000000000000008;
  cdoDSNSuccessFailOrDelay = $000000000000000E;
Type
  CdoEventStatus =LongWord;
Const
  cdoRunNextSink = $0000000000000000;
  cdoSkipRemainingSinks = $0000000000000001;
Type
  cdoImportanceValues =LongWord;
Const
  cdoLow = $0000000000000000;
  cdoNormal = $0000000000000001;
  cdoHigh = $0000000000000002;
Type
  CdoMessageStat =LongWord;
Const
  cdoStatSuccess = $0000000000000000;
  cdoStatAbortDelivery = $0000000000000002;
  cdoStatBadMail = $0000000000000003;
Type
  CdoMHTMLFlags =LongWord;
Const
  cdoSuppressNone = $0000000000000000;
  cdoSuppressImages = $0000000000000001;
  cdoSuppressBGSounds = $0000000000000002;
  cdoSuppressFrames = $0000000000000004;
  cdoSuppressObjects = $0000000000000008;
  cdoSuppressStyleSheets = $0000000000000010;
  cdoSuppressAll = $000000000000001F;
Type
  CdoNNTPProcessingField =LongWord;
Const
  cdoPostMessage = $0000000000000001;
  cdoProcessControl = $0000000000000002;
  cdoProcessModerator = $0000000000000004;
Type
  CdoPostUsing =LongWord;
Const
  cdoPostUsingPickup = $0000000000000001;
  cdoPostUsingPort = $0000000000000002;
Type
  cdoPriorityValues =LongWord;
Const
  cdoPriorityNonUrgent = $FFFFFFFFFFFFFFFF;
  cdoPriorityNormal = $0000000000000000;
  cdoPriorityUrgent = $0000000000000001;
Type
  CdoProtocolsAuthentication =LongWord;
Const
  cdoAnonymous = $0000000000000000;
  cdoBasic = $0000000000000001;
  cdoNTLM = $0000000000000002;
Type
  CdoReferenceType =LongWord;
Const
  cdoRefTypeId = $0000000000000000;
  cdoRefTypeLocation = $0000000000000001;
Type
  CdoSendUsing =LongWord;
Const
  cdoSendUsingPickup = $0000000000000001;
  cdoSendUsingPort = $0000000000000002;
Type
  cdoSensitivityValues =LongWord;
Const
  cdoSensitivityNone = $0000000000000000;
  cdoPersonal = $0000000000000001;
  cdoPrivate = $0000000000000002;
  cdoCompanyConfidential = $0000000000000003;
Type
  CdoTimeZoneId =LongWord;
Const
  cdoUTC = $0000000000000000;
  cdoGMT = $0000000000000001;
  cdoSarajevo = $0000000000000002;
  cdoParis = $0000000000000003;
  cdoBerlin = $0000000000000004;
  cdoEasternEurope = $0000000000000005;
  cdoPrague = $0000000000000006;
  cdoAthens = $0000000000000007;
  cdoBrasilia = $0000000000000008;
  cdoAtlanticCanada = $0000000000000009;
  cdoEastern = $000000000000000A;
  cdoCentral = $000000000000000B;
  cdoMountain = $000000000000000C;
  cdoPacific = $000000000000000D;
  cdoAlaska = $000000000000000E;
  cdoHawaii = $000000000000000F;
  cdoMidwayIsland = $0000000000000010;
  cdoWellington = $0000000000000011;
  cdoBrisbane = $0000000000000012;
  cdoAdelaide = $0000000000000013;
  cdoTokyo = $0000000000000014;
  cdoSingapore = $0000000000000015;
  cdoBangkok = $0000000000000016;
  cdoBombay = $0000000000000017;
  cdoAbuDhabi = $0000000000000018;
  cdoTehran = $0000000000000019;
  cdoBaghdad = $000000000000001A;
  cdoIsrael = $000000000000001B;
  cdoNewfoundland = $000000000000001C;
  cdoAzores = $000000000000001D;
  cdoMidAtlantic = $000000000000001E;
  cdoMonrovia = $000000000000001F;
  cdoBuenosAires = $0000000000000020;
  cdoCaracas = $0000000000000021;
  cdoIndiana = $0000000000000022;
  cdoBogota = $0000000000000023;
  cdoSaskatchewan = $0000000000000024;
  cdoMexicoCity = $0000000000000025;
  cdoArizona = $0000000000000026;
  cdoEniwetok = $0000000000000027;
  cdoFiji = $0000000000000028;
  cdoMagadan = $0000000000000029;
  cdoHobart = $000000000000002A;
  cdoGuam = $000000000000002B;
  cdoDarwin = $000000000000002C;
  cdoBeijing = $000000000000002D;
  cdoAlmaty = $000000000000002E;
  cdoIslamabad = $000000000000002F;
  cdoKabul = $0000000000000030;
  cdoCairo = $0000000000000031;
  cdoHarare = $0000000000000032;
  cdoMoscow = $0000000000000033;
  cdoFloating = $0000000000000034;
  cdoCapeVerde = $0000000000000035;
  cdoCaucasus = $0000000000000036;
  cdoCentralAmerica = $0000000000000037;
  cdoEastAfrica = $0000000000000038;
  cdoMelbourne = $0000000000000039;
  cdoEkaterinburg = $000000000000003A;
  cdoHelsinki = $000000000000003B;
  cdoGreenland = $000000000000003C;
  cdoRangoon = $000000000000003D;
  cdoNepal = $000000000000003E;
  cdoIrkutsk = $000000000000003F;
  cdoKrasnoyarsk = $0000000000000040;
  cdoSantiago = $0000000000000041;
  cdoSriLanka = $0000000000000042;
  cdoTonga = $0000000000000043;
  cdoVladivostok = $0000000000000044;
  cdoWestCentralAfrica = $0000000000000045;
  cdoYakutsk = $0000000000000046;
  cdoDhaka = $0000000000000047;
  cdoSeoul = $0000000000000048;
  cdoPerth = $0000000000000049;
  cdoArab = $000000000000004A;
  cdoTaipei = $000000000000004B;
  cdoSydney2000 = $000000000000004C;
  cdoChihuahua = $000000000000004D;
  cdoCanberraCommonwealthGames2006 = $000000000000004E;
  cdoAdelaideCommonwealthGames2006 = $000000000000004F;
  cdoHobartCommonwealthGames2006 = $0000000000000050;
  cdoTijuana = $0000000000000051;
  cdoInvalidTimeZone = $0000000000000052;
//Forward declarations

Type
 IBodyParts = interface;
 IBodyPartsDisp = dispinterface;
 IBodyPart = interface;
 IBodyPartDisp = dispinterface;
 IDataSource = interface;
 IDataSourceDisp = dispinterface;
 IMessages = interface;
 IMessagesDisp = dispinterface;
 IMessage = interface;
 IMessageDisp = dispinterface;
 IConfiguration = interface;
 IConfigurationDisp = dispinterface;
 IDropDirectory = interface;
 IDropDirectoryDisp = dispinterface;
 ISMTPScriptConnector = interface;
 ISMTPScriptConnectorDisp = dispinterface;
 ISMTPOnArrival = interface;
 ISMTPOnArrivalDisp = dispinterface;
 INNTPEarlyScriptConnector = interface;
 INNTPEarlyScriptConnectorDisp = dispinterface;
 INNTPOnPostEarly = interface;
 INNTPOnPostEarlyDisp = dispinterface;
 INNTPPostScriptConnector = interface;
 INNTPPostScriptConnectorDisp = dispinterface;
 INNTPOnPost = interface;
 INNTPOnPostDisp = dispinterface;
 INNTPFinalScriptConnector = interface;
 INNTPFinalScriptConnectorDisp = dispinterface;
 INNTPOnPostFinal = interface;
 INNTPOnPostFinalDisp = dispinterface;

//Map CoClass to its default interface

 Message = IMessage;
 Configuration = IConfiguration;
 DropDirectory = IDropDirectory;
 SMTPConnector = ISMTPScriptConnector;
 NNTPEarlyConnector = INNTPEarlyScriptConnector;
 NNTPPostConnector = INNTPPostScriptConnector;
 NNTPFinalConnector = INNTPFinalScriptConnector;

//records, unions, aliases


//interface declarations

// IBodyParts : Defines methods and properties used to manage a collection of BodyPart objects.

 IBodyParts = interface(IDispatch)
   ['{CD000023-8B95-11D1-82DB-00C04FB1625D}']
   function Get_Count : Integer; safecall;
   function Get_Item(Index:Integer) : IBodyPart; safecall;
   function Get__NewEnum : IUnknown; safecall;
    // Delete : Deletes the specified BodyPart object from the collection. Can use the index or a reference to the object. 
   procedure Delete(varBP:OleVariant);safecall;
    // DeleteAll : Deletes all BodyPart objects in the collection. 
   procedure DeleteAll;safecall;
    // Add : Adds a BodyPart object to the collection at the specified index, and returns the newly added object. 
   function Add(Index:Integer):IBodyPart;safecall;
    // Count : Returns the number of BodyPart objects in the collection. 
   property Count:Integer read Get_Count;
    // Item : The specified BodyPart object in the collection. 
   property Item[Index:Integer]:IBodyPart read Get_Item; default;
    // _NewEnum :  
   property _NewEnum:IUnknown read Get__NewEnum;
  end;


// IBodyParts : Defines methods and properties used to manage a collection of BodyPart objects.

 IBodyPartsDisp = dispinterface
   ['{CD000023-8B95-11D1-82DB-00C04FB1625D}']
    // Delete : Deletes the specified BodyPart object from the collection. Can use the index or a reference to the object. 
   procedure Delete(varBP:OleVariant);dispid 2;
    // DeleteAll : Deletes all BodyPart objects in the collection. 
   procedure DeleteAll;dispid 3;
    // Add : Adds a BodyPart object to the collection at the specified index, and returns the newly added object. 
   function Add(Index:Integer):IBodyPart;dispid 4;
    // Count : Returns the number of BodyPart objects in the collection. 
   property Count:Integer  readonly dispid 1;
    // Item : The specified BodyPart object in the collection. 
   property Item[Index:Integer]:IBodyPart  readonly dispid 0; default;
    // _NewEnum :  
   property _NewEnum:IUnknown  readonly dispid -4;
  end;


// IBodyPart : Defines methods, properties, and collections used to manage a message body part.

 IBodyPart = interface(IDispatch)
   ['{CD000021-8B95-11D1-82DB-00C04FB1625D}']
   function Get_BodyParts : IBodyParts; safecall;
   function Get_ContentTransferEncoding : WideString; safecall;
   procedure Set_ContentTransferEncoding(const pContentTransferEncoding:WideString); safecall;
   function Get_ContentMediaType : WideString; safecall;
   procedure Set_ContentMediaType(const pContentMediaType:WideString); safecall;
   function Get_Fields : Fields; safecall;
   function Get_Charset : WideString; safecall;
   procedure Set_Charset(const pCharset:WideString); safecall;
   function Get_FileName : WideString; safecall;
   function Get_DataSource : IDataSource; safecall;
   function Get_ContentClass : WideString; safecall;
   procedure Set_ContentClass(const pContentClass:WideString); safecall;
   function Get_ContentClassName : WideString; safecall;
   procedure Set_ContentClassName(const pContentClassName:WideString); safecall;
   function Get_Parent : IBodyPart; safecall;
    // AddBodyPart : Adds a body part to the object's BodyParts collection. 
   function AddBodyPart(Index:Integer):IBodyPart;safecall;
    // SaveToFile : Saves the body part content to the specified file. 
   procedure SaveToFile(FileName:WideString);safecall;
    // GetEncodedContentStream : Returns a Stream object containing the body part content in encoded format. The encoding method is specified in the ContentTransferEncoding property. 
   function GetEncodedContentStream:_Stream;safecall;
    // GetDecodedContentStream : Returns a Stream object containing the body part content in decoded format. 
   function GetDecodedContentStream:_Stream;safecall;
    // GetStream : Returns an ADO Stream object containing the body part in serialized, MIME encoded format. 
   function GetStream:_Stream;safecall;
    // GetFieldParameter : Returns the specified parameter from the body part's specified header field. 
   function GetFieldParameter(FieldName:WideString;Parameter:WideString):WideString;safecall;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;safecall;
    // BodyParts : The object's BodyParts collection. 
   property BodyParts:IBodyParts read Get_BodyParts;
    // ContentTransferEncoding : The method used to encode the body part content. For example, quoted-printable or base64. 
   property ContentTransferEncoding:WideString read Get_ContentTransferEncoding write Set_ContentTransferEncoding;
    // ContentMediaType : The content media type portion of the body part's content type. 
   property ContentMediaType:WideString read Get_ContentMediaType write Set_ContentMediaType;
    // Fields : The object's Fields collection. 
   property Fields:Fields read Get_Fields;
    // Charset : The character set of the body part's text content (not applicable for non-text content types). 
   property Charset:WideString read Get_Charset write Set_Charset;
    // FileName : The value of the filename parameter for the content-disposition MIME header. 
   property FileName:WideString read Get_FileName;
    // DataSource : The object's IDataSource interface. 
   property DataSource:IDataSource read Get_DataSource;
    // ContentClass : The body part's content class. 
   property ContentClass:WideString read Get_ContentClass write Set_ContentClass;
    // ContentClassName : Deprecated. Do not use. 
   property ContentClassName:WideString read Get_ContentClassName write Set_ContentClassName;
    // Parent : The object's parent object in the body part hierarchy. 
   property Parent:IBodyPart read Get_Parent;
  end;


// IBodyPart : Defines methods, properties, and collections used to manage a message body part.

 IBodyPartDisp = dispinterface
   ['{CD000021-8B95-11D1-82DB-00C04FB1625D}']
    // AddBodyPart : Adds a body part to the object's BodyParts collection. 
   function AddBodyPart(Index:Integer):IBodyPart;dispid 250;
    // SaveToFile : Saves the body part content to the specified file. 
   procedure SaveToFile(FileName:WideString);dispid 251;
    // GetEncodedContentStream : Returns a Stream object containing the body part content in encoded format. The encoding method is specified in the ContentTransferEncoding property. 
   function GetEncodedContentStream:_Stream;dispid 252;
    // GetDecodedContentStream : Returns a Stream object containing the body part content in decoded format. 
   function GetDecodedContentStream:_Stream;dispid 253;
    // GetStream : Returns an ADO Stream object containing the body part in serialized, MIME encoded format. 
   function GetStream:_Stream;dispid 254;
    // GetFieldParameter : Returns the specified parameter from the body part's specified header field. 
   function GetFieldParameter(FieldName:WideString;Parameter:WideString):WideString;dispid 255;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;dispid 160;
    // BodyParts : The object's BodyParts collection. 
   property BodyParts:IBodyParts  readonly dispid 200;
    // ContentTransferEncoding : The method used to encode the body part content. For example, quoted-printable or base64. 
   property ContentTransferEncoding:WideString dispid 201;
    // ContentMediaType : The content media type portion of the body part's content type. 
   property ContentMediaType:WideString dispid 202;
    // Fields : The object's Fields collection. 
   property Fields:Fields  readonly dispid 203;
    // Charset : The character set of the body part's text content (not applicable for non-text content types). 
   property Charset:WideString dispid 204;
    // FileName : The value of the filename parameter for the content-disposition MIME header. 
   property FileName:WideString  readonly dispid 205;
    // DataSource : The object's IDataSource interface. 
   property DataSource:IDataSource  readonly dispid 207;
    // ContentClass : The body part's content class. 
   property ContentClass:WideString dispid 208;
    // ContentClassName : Deprecated. Do not use. 
   property ContentClassName:WideString dispid 209;
    // Parent : The object's parent object in the body part hierarchy. 
   property Parent:IBodyPart  readonly dispid 210;
  end;


// IDataSource : Defines methods, properties, and collections used to extract messages from or embed messages into other CDO message body parts.

 IDataSource = interface(IDispatch)
   ['{CD000029-8B95-11D1-82DB-00C04FB1625D}']
   function Get_SourceClass : WideString; safecall;
   function Get_Source : IUnknown; safecall;
   function Get_IsDirty : WordBool; safecall;
   procedure Set_IsDirty(const pIsDirty:WordBool); safecall;
   function Get_SourceURL : WideString; safecall;
   function Get_ActiveConnection : _Connection; safecall;
    // SaveToObject : Binds to and saves data into the specified object. 
   procedure SaveToObject(Source:IUnknown;InterfaceName:WideString);safecall;
    // OpenObject : Binds to and opens data from the specified object. 
   procedure OpenObject(Source:IUnknown;InterfaceName:WideString);safecall;
    // SaveTo : Not implemented. Reserved for future use. 
   procedure SaveTo(SourceURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);safecall;
    // Open : Not implemented. Reserved for future use. 
   procedure Open(SourceURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);safecall;
    // Save : Saves data into the currently bound object. 
   procedure Save;safecall;
    // SaveToContainer : Not implemented. Reserved for future use. 
   procedure SaveToContainer(ContainerURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);safecall;
    // SourceClass : The interface name (type) of the currently bound object. When you bind resources by URL, the value _Record is returned. 
   property SourceClass:WideString read Get_SourceClass;
    // Source : Returns the currently bound object. When you bind resources by URL, an ADO _Record interface is returned on an open Record object. 
   property Source:IUnknown read Get_Source;
    // IsDirty : Indicates whether the local data has been changed since the last save or bind operation. 
   property IsDirty:WordBool read Get_IsDirty write Set_IsDirty;
    // SourceURL : Not Implemented. Reserved for future use. 
   property SourceURL:WideString read Get_SourceURL;
    // ActiveConnection : Not implemented. Reserved for future use. 
   property ActiveConnection:_Connection read Get_ActiveConnection;
  end;


// IDataSource : Defines methods, properties, and collections used to extract messages from or embed messages into other CDO message body parts.

 IDataSourceDisp = dispinterface
   ['{CD000029-8B95-11D1-82DB-00C04FB1625D}']
    // SaveToObject : Binds to and saves data into the specified object. 
   procedure SaveToObject(Source:IUnknown;InterfaceName:WideString);dispid 251;
    // OpenObject : Binds to and opens data from the specified object. 
   procedure OpenObject(Source:IUnknown;InterfaceName:WideString);dispid 252;
    // SaveTo : Not implemented. Reserved for future use. 
   procedure SaveTo(SourceURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);dispid 253;
    // Open : Not implemented. Reserved for future use. 
   procedure Open(SourceURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);dispid 254;
    // Save : Saves data into the currently bound object. 
   procedure Save;dispid 255;
    // SaveToContainer : Not implemented. Reserved for future use. 
   procedure SaveToContainer(ContainerURL:WideString;ActiveConnection:IDispatch;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);dispid 256;
    // SourceClass : The interface name (type) of the currently bound object. When you bind resources by URL, the value _Record is returned. 
   property SourceClass:WideString  readonly dispid 207;
    // Source : Returns the currently bound object. When you bind resources by URL, an ADO _Record interface is returned on an open Record object. 
   property Source:IUnknown  readonly dispid 208;
    // IsDirty : Indicates whether the local data has been changed since the last save or bind operation. 
   property IsDirty:WordBool dispid 209;
    // SourceURL : Not Implemented. Reserved for future use. 
   property SourceURL:WideString  readonly dispid 210;
    // ActiveConnection : Not implemented. Reserved for future use. 
   property ActiveConnection:_Connection  readonly dispid 211;
  end;


// IMessages : Defines methods and properties used to manage a collection of Message objects on the file system. Returned by IDropDirectory.GetMessages.

 IMessages = interface(IDispatch)
   ['{CD000025-8B95-11D1-82DB-00C04FB1625D}']
   function Get_Item : IMessage; safecall;
   function Get_Count : Integer; safecall;
    // Delete : Deletes the specified message object in the collection. 
   procedure Delete(Index:Integer);safecall;
    // DeleteAll : Deletes all message objects in the collection. 
   procedure DeleteAll;safecall;
   function Get__NewEnum : IUnknown; safecall;
   function Get_FileName : WideString; safecall;
    // Item : Returns the message specified by index from the collection. 
   property Item:IMessage read Get_Item;
    // Count : The number of message objects in the collection. 
   property Count:Integer read Get_Count;
    // _NewEnum :  
   property _NewEnum:IUnknown read Get__NewEnum;
    // FileName : Returns the name of the file containing the specified message in the file system. 
   property FileName:WideString read Get_FileName;
  end;


// IMessages : Defines methods and properties used to manage a collection of Message objects on the file system. Returned by IDropDirectory.GetMessages.

 IMessagesDisp = dispinterface
   ['{CD000025-8B95-11D1-82DB-00C04FB1625D}']
    // Delete : Deletes the specified message object in the collection. 
   procedure Delete(Index:Integer);dispid 2;
    // DeleteAll : Deletes all message objects in the collection. 
   procedure DeleteAll;dispid 3;
    // Item : Returns the message specified by index from the collection. 
   property Item:IMessage  readonly dispid 0;
    // Count : The number of message objects in the collection. 
   property Count:Integer  readonly dispid 1;
    // _NewEnum :  
   property _NewEnum:IUnknown  readonly dispid -4;
    // FileName : Returns the name of the file containing the specified message in the file system. 
   property FileName:WideString  readonly dispid 5;
  end;


// IMessage : Defines methods, properties, and collections used to manage a message.

 IMessage = interface(IDispatch)
   ['{CD000020-8B95-11D1-82DB-00C04FB1625D}']
   function Get_BCC : WideString; safecall;
   procedure Set_BCC(const pBCC:WideString); safecall;
   function Get_CC : WideString; safecall;
   procedure Set_CC(const pCC:WideString); safecall;
   function Get_FollowUpTo : WideString; safecall;
   procedure Set_FollowUpTo(const pFollowUpTo:WideString); safecall;
   function Get_From : WideString; safecall;
   procedure Set_From(const pFrom:WideString); safecall;
   function Get_Keywords : WideString; safecall;
   procedure Set_Keywords(const pKeywords:WideString); safecall;
   function Get_MimeFormatted : WordBool; safecall;
   procedure Set_MimeFormatted(const pMimeFormatted:WordBool); safecall;
   function Get_Newsgroups : WideString; safecall;
   procedure Set_Newsgroups(const pNewsgroups:WideString); safecall;
   function Get_Organization : WideString; safecall;
   procedure Set_Organization(const pOrganization:WideString); safecall;
   function Get_ReceivedTime : TDateTime; safecall;
   function Get_ReplyTo : WideString; safecall;
   procedure Set_ReplyTo(const pReplyTo:WideString); safecall;
   function Get_DSNOptions : CdoDSNOptions; safecall;
   procedure Set_DSNOptions(const pDSNOptions:CdoDSNOptions); safecall;
   function Get_SentOn : TDateTime; safecall;
   function Get_Subject : WideString; safecall;
   procedure Set_Subject(const pSubject:WideString); safecall;
   function Get_To_ : WideString; safecall;
   procedure Set_To_(const pTo:WideString); safecall;
   function Get_TextBody : WideString; safecall;
   procedure Set_TextBody(const pTextBody:WideString); safecall;
   function Get_HTMLBody : WideString; safecall;
   procedure Set_HTMLBody(const pHTMLBody:WideString); safecall;
   function Get_Attachments : IBodyParts; safecall;
   function Get_Sender : WideString; safecall;
   procedure Set_Sender(const pSender:WideString); safecall;
   function Get_Configuration : IConfiguration; safecall;
   procedure Set_Configuration(const pConfiguration:IConfiguration); safecall;
   procedure Set_Configuration_(const pConfiguration:IConfiguration); safecall;
   function Get_AutoGenerateTextBody : WordBool; safecall;
   procedure Set_AutoGenerateTextBody(const pAutoGenerateTextBody:WordBool); safecall;
   function Get_EnvelopeFields : Fields; safecall;
   function Get_TextBodyPart : IBodyPart; safecall;
   function Get_HTMLBodyPart : IBodyPart; safecall;
   function Get_BodyPart : IBodyPart; safecall;
   function Get_DataSource : IDataSource; safecall;
   function Get_Fields : Fields; safecall;
   function Get_MDNRequested : WordBool; safecall;
   procedure Set_MDNRequested(const pMDNRequested:WordBool); safecall;
    // AddRelatedBodyPart : Adds a BodyPart object with content referenced within the text/html portion of the message body. 
   function AddRelatedBodyPart(URL:WideString;Reference:WideString;ReferenceType:CdoReferenceType;UserName:WideString;Password:WideString):IBodyPart;safecall;
    // AddAttachment : Adds an attachment (BodyPart) to the message. 
   function AddAttachment(URL:WideString;UserName:WideString;Password:WideString):IBodyPart;safecall;
    // CreateMHTMLBody : Creates an MHTML-formatted message body using the resource(s) at the specified URL. 
   procedure CreateMHTMLBody(URL:WideString;Flags:CdoMHTMLFlags;UserName:WideString;Password:WideString);safecall;
    // Forward : Returns a Message object used to forward a message. 
   function Forward:IMessage;safecall;
    // Post : Posts the message using the method specified in the associated Configuration object. 
   procedure Post;safecall;
    // PostReply : Returns a Message object used to post a reply to the message. 
   function PostReply:IMessage;safecall;
    // Reply : Returns a Message object used to reply to the message. 
   function Reply:IMessage;safecall;
    // ReplyAll : Returns a Message object used to post a reply to all recipients of the message. 
   function ReplyAll:IMessage;safecall;
    // Send : Sends the message using the method specified in the associated Configuration object. 
   procedure Send;safecall;
    // GetStream : Returns an ADO Stream object containing the message in serialized, RFC 822 format. The message body is encoded using either MIME or UUENCODE as specified by the MIMEFormatted property. 
   function GetStream:_Stream;safecall;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;safecall;
    // BCC : The message's hidden carbon copy (BCC header) recipients. 
   property BCC:WideString read Get_BCC write Set_BCC;
    // CC : The message's secondary (CC header) recipients. 
   property CC:WideString read Get_CC write Set_CC;
    // FollowUpTo : The message's follow-up recipients. 
   property FollowUpTo:WideString read Get_FollowUpTo write Set_FollowUpTo;
    // From : The message's principle (From header) authors. 
   property From:WideString read Get_From write Set_From;
    // Keywords : The message's keywords. 
   property Keywords:WideString read Get_Keywords write Set_Keywords;
    // MimeFormatted : Indicates whether the message is to be serialized using the MIME (True) or UUENCODE (False) format. 
   property MimeFormatted:WordBool read Get_MimeFormatted write Set_MimeFormatted;
    // Newsgroups : The message's newsgroup (Newsgroups header) recipients. 
   property Newsgroups:WideString read Get_Newsgroups write Set_Newsgroups;
    // Organization : The sender's organization name. 
   property Organization:WideString read Get_Organization write Set_Organization;
    // ReceivedTime : The date and time the message was received. 
   property ReceivedTime:TDateTime read Get_ReceivedTime;
    // ReplyTo : The email addresses (Reply-To header) to which to reply. 
   property ReplyTo:WideString read Get_ReplyTo write Set_ReplyTo;
    // DSNOptions : The delivery status notification (DSN) options for the message. 
   property DSNOptions:CdoDSNOptions read Get_DSNOptions write Set_DSNOptions;
    // SentOn : The date and time the message was sent. 
   property SentOn:TDateTime read Get_SentOn;
    // Subject : The message's subject (Subject header). 
   property Subject:WideString read Get_Subject write Set_Subject;
    // To : The message's principle (To header) recipients. 
   property To_:WideString read Get_To_ write Set_To_;
    // TextBody : The text/plain portion of the message body. 
   property TextBody:WideString read Get_TextBody write Set_TextBody;
    // HTMLBody : The text/html portion of the message body. 
   property HTMLBody:WideString read Get_HTMLBody write Set_HTMLBody;
    // Attachments : The object's Attachments collection. 
   property Attachments:IBodyParts read Get_Attachments;
    // Sender : The message's actual sender. 
   property Sender:WideString read Get_Sender write Set_Sender;
    // Configuration : The object's Configuration object. 
   property Configuration:IConfiguration read Get_Configuration write Set_Configuration;
    // AutoGenerateTextBody : Indicates whether a text/plain alternate representation should automatically be generated from the text/html part of the message body. 
   property AutoGenerateTextBody:WordBool read Get_AutoGenerateTextBody write Set_AutoGenerateTextBody;
    // EnvelopeFields : The transport envelope Fields collection for the message (transport event sinks only). 
   property EnvelopeFields:Fields read Get_EnvelopeFields;
    // TextBodyPart : Returns the BodyPart object (IBodyPart interface) containing the text/plain part of the message body. 
   property TextBodyPart:IBodyPart read Get_TextBodyPart;
    // HTMLBodyPart : Returns the BodyPart object (IBodyPart interface) containing the text/html portion of the message body. 
   property HTMLBodyPart:IBodyPart read Get_HTMLBodyPart;
    // BodyPart : The object's IBodyPart interface. 
   property BodyPart:IBodyPart read Get_BodyPart;
    // DataSource : The object's IDataSource interface. 
   property DataSource:IDataSource read Get_DataSource;
    // Fields : The object's Fields collection. 
   property Fields:Fields read Get_Fields;
    // MDNRequested : Indicates whether a mail delivery notification (MDN) should be sent when the message is received. 
   property MDNRequested:WordBool read Get_MDNRequested write Set_MDNRequested;
  end;


// IMessage : Defines methods, properties, and collections used to manage a message.

 IMessageDisp = dispinterface
   ['{CD000020-8B95-11D1-82DB-00C04FB1625D}']
    // AddRelatedBodyPart : Adds a BodyPart object with content referenced within the text/html portion of the message body. 
   function AddRelatedBodyPart(URL:WideString;Reference:WideString;ReferenceType:CdoReferenceType;UserName:WideString;Password:WideString):IBodyPart;dispid 150;
    // AddAttachment : Adds an attachment (BodyPart) to the message. 
   function AddAttachment(URL:WideString;UserName:WideString;Password:WideString):IBodyPart;dispid 151;
    // CreateMHTMLBody : Creates an MHTML-formatted message body using the resource(s) at the specified URL. 
   procedure CreateMHTMLBody(URL:WideString;Flags:CdoMHTMLFlags;UserName:WideString;Password:WideString);dispid 152;
    // Forward : Returns a Message object used to forward a message. 
   function Forward:IMessage;dispid 153;
    // Post : Posts the message using the method specified in the associated Configuration object. 
   procedure Post;dispid 154;
    // PostReply : Returns a Message object used to post a reply to the message. 
   function PostReply:IMessage;dispid 155;
    // Reply : Returns a Message object used to reply to the message. 
   function Reply:IMessage;dispid 156;
    // ReplyAll : Returns a Message object used to post a reply to all recipients of the message. 
   function ReplyAll:IMessage;dispid 157;
    // Send : Sends the message using the method specified in the associated Configuration object. 
   procedure Send;dispid 158;
    // GetStream : Returns an ADO Stream object containing the message in serialized, RFC 822 format. The message body is encoded using either MIME or UUENCODE as specified by the MIMEFormatted property. 
   function GetStream:_Stream;dispid 159;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;dispid 160;
    // BCC : The message's hidden carbon copy (BCC header) recipients. 
   property BCC:WideString dispid 101;
    // CC : The message's secondary (CC header) recipients. 
   property CC:WideString dispid 103;
    // FollowUpTo : The message's follow-up recipients. 
   property FollowUpTo:WideString dispid 105;
    // From : The message's principle (From header) authors. 
   property From:WideString dispid 106;
    // Keywords : The message's keywords. 
   property Keywords:WideString dispid 107;
    // MimeFormatted : Indicates whether the message is to be serialized using the MIME (True) or UUENCODE (False) format. 
   property MimeFormatted:WordBool dispid 110;
    // Newsgroups : The message's newsgroup (Newsgroups header) recipients. 
   property Newsgroups:WideString dispid 111;
    // Organization : The sender's organization name. 
   property Organization:WideString dispid 112;
    // ReceivedTime : The date and time the message was received. 
   property ReceivedTime:TDateTime  readonly dispid 114;
    // ReplyTo : The email addresses (Reply-To header) to which to reply. 
   property ReplyTo:WideString dispid 115;
    // DSNOptions : The delivery status notification (DSN) options for the message. 
   property DSNOptions:CdoDSNOptions dispid 116;
    // SentOn : The date and time the message was sent. 
   property SentOn:TDateTime  readonly dispid 119;
    // Subject : The message's subject (Subject header). 
   property Subject:WideString dispid 120;
    // To : The message's principle (To header) recipients. 
   property To_:WideString dispid 121;
    // TextBody : The text/plain portion of the message body. 
   property TextBody:WideString dispid 123;
    // HTMLBody : The text/html portion of the message body. 
   property HTMLBody:WideString dispid 124;
    // Attachments : The object's Attachments collection. 
   property Attachments:IBodyParts  readonly dispid 125;
    // Sender : The message's actual sender. 
   property Sender:WideString dispid 126;
    // Configuration : The object's Configuration object. 
   property Configuration:IConfiguration dispid 127;
    // AutoGenerateTextBody : Indicates whether a text/plain alternate representation should automatically be generated from the text/html part of the message body. 
   property AutoGenerateTextBody:WordBool dispid 128;
    // EnvelopeFields : The transport envelope Fields collection for the message (transport event sinks only). 
   property EnvelopeFields:Fields  readonly dispid 129;
    // TextBodyPart : Returns the BodyPart object (IBodyPart interface) containing the text/plain part of the message body. 
   property TextBodyPart:IBodyPart  readonly dispid 130;
    // HTMLBodyPart : Returns the BodyPart object (IBodyPart interface) containing the text/html portion of the message body. 
   property HTMLBodyPart:IBodyPart  readonly dispid 131;
    // BodyPart : The object's IBodyPart interface. 
   property BodyPart:IBodyPart  readonly dispid 132;
    // DataSource : The object's IDataSource interface. 
   property DataSource:IDataSource  readonly dispid 133;
    // Fields : The object's Fields collection. 
   property Fields:Fields  readonly dispid 134;
    // MDNRequested : Indicates whether a mail delivery notification (MDN) should be sent when the message is received. 
   property MDNRequested:WordBool dispid 135;
  end;


// IConfiguration : Defines methods, properties, and collections used to manage configuration information for CDO objects.

 IConfiguration = interface(IDispatch)
   ['{CD000022-8B95-11D1-82DB-00C04FB1625D}']
   function Get_Fields : Fields; safecall;
    // Load : Loads the specified configuration. 
   procedure Load(LoadFrom:CdoConfigSource;URL:WideString);safecall;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;safecall;
    // Fields : The object's Fields collection. 
   property Fields:Fields read Get_Fields;
  end;


// IConfiguration : Defines methods, properties, and collections used to manage configuration information for CDO objects.

 IConfigurationDisp = dispinterface
   ['{CD000022-8B95-11D1-82DB-00C04FB1625D}']
    // Load : Loads the specified configuration. 
   procedure Load(LoadFrom:CdoConfigSource;URL:WideString);dispid 50;
    // GetInterface : Returns a specified interface on this object; provided for script languages. 
   function GetInterface(Interface_:WideString):IDispatch;dispid 160;
    // Fields : The object's Fields collection. 
   property Fields:Fields  readonly dispid 0;
  end;


// IDropDirectory : Defines methods, properties, and collections used to manage a collection of messages on the file system.

 IDropDirectory = interface(IDispatch)
   ['{CD000024-8B95-11D1-82DB-00C04FB1625D}']
    // GetMessages : Returns a collection of messages contained in the specified directory on the file system. The default location is the SMTP drop directory. 
   function GetMessages(DirName:WideString):IMessages;safecall;
  end;


// IDropDirectory : Defines methods, properties, and collections used to manage a collection of messages on the file system.

 IDropDirectoryDisp = dispinterface
   ['{CD000024-8B95-11D1-82DB-00C04FB1625D}']
    // GetMessages : Returns a collection of messages contained in the specified directory on the file system. The default location is the SMTP drop directory. 
   function GetMessages(DirName:WideString):IMessages;dispid 200;
  end;


// ISMTPScriptConnector : ISMTPScriptConnector interface

 ISMTPScriptConnector = interface(IDispatch)
   ['{CD000030-8B95-11D1-82DB-00C04FB1625D}']
  end;


// ISMTPScriptConnector : ISMTPScriptConnector interface

 ISMTPScriptConnectorDisp = dispinterface
   ['{CD000030-8B95-11D1-82DB-00C04FB1625D}']
  end;


// ISMTPOnArrival : Implement when creating SMTP OnArrival event sinks.

 ISMTPOnArrival = interface(IDispatch)
   ['{CD000026-8B95-11D1-82DB-00C04FB1625D}']
    // OnArrival : Called by the SMTP event dispatcher when a message arrives. 
   procedure OnArrival(Msg:IMessage;var EventStatus:CdoEventStatus);safecall;
  end;


// ISMTPOnArrival : Implement when creating SMTP OnArrival event sinks.

 ISMTPOnArrivalDisp = dispinterface
   ['{CD000026-8B95-11D1-82DB-00C04FB1625D}']
    // OnArrival : Called by the SMTP event dispatcher when a message arrives. 
   procedure OnArrival(Msg:IMessage;var EventStatus:CdoEventStatus);dispid 256;
  end;


// INNTPEarlyScriptConnector : INNTPFinalScriptConnector interface

 INNTPEarlyScriptConnector = interface(IDispatch)
   ['{CD000034-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPEarlyScriptConnector : INNTPFinalScriptConnector interface

 INNTPEarlyScriptConnectorDisp = dispinterface
   ['{CD000034-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPOnPostEarly : Implement when creating NNTP OnPostEarly event sinks.

 INNTPOnPostEarly = interface(IDispatch)
   ['{CD000033-8B95-11D1-82DB-00C04FB1625D}']
    // OnPostEarly : Called by the NNTP event dispatcher when message headers arrive. 
   procedure OnPostEarly(Msg:IMessage;var EventStatus:CdoEventStatus);safecall;
  end;


// INNTPOnPostEarly : Implement when creating NNTP OnPostEarly event sinks.

 INNTPOnPostEarlyDisp = dispinterface
   ['{CD000033-8B95-11D1-82DB-00C04FB1625D}']
    // OnPostEarly : Called by the NNTP event dispatcher when message headers arrive. 
   procedure OnPostEarly(Msg:IMessage;var EventStatus:CdoEventStatus);dispid 256;
  end;


// INNTPPostScriptConnector : INNTPPostScriptConnector interface

 INNTPPostScriptConnector = interface(IDispatch)
   ['{CD000031-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPPostScriptConnector : INNTPPostScriptConnector interface

 INNTPPostScriptConnectorDisp = dispinterface
   ['{CD000031-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPOnPost : Implement when creating NNTP OnPost event sinks.

 INNTPOnPost = interface(IDispatch)
   ['{CD000027-8B95-11D1-82DB-00C04FB1625D}']
    // OnPost : Called by the NNTP event dispatcher when a message is posted. 
   procedure OnPost(Msg:IMessage;var EventStatus:CdoEventStatus);safecall;
  end;


// INNTPOnPost : Implement when creating NNTP OnPost event sinks.

 INNTPOnPostDisp = dispinterface
   ['{CD000027-8B95-11D1-82DB-00C04FB1625D}']
    // OnPost : Called by the NNTP event dispatcher when a message is posted. 
   procedure OnPost(Msg:IMessage;var EventStatus:CdoEventStatus);dispid 256;
  end;


// INNTPFinalScriptConnector : INNTPFinalScriptConnector interface

 INNTPFinalScriptConnector = interface(IDispatch)
   ['{CD000032-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPFinalScriptConnector : INNTPFinalScriptConnector interface

 INNTPFinalScriptConnectorDisp = dispinterface
   ['{CD000032-8B95-11D1-82DB-00C04FB1625D}']
  end;


// INNTPOnPostFinal : Implement when creating NNTP OnPostFinal event sinks.

 INNTPOnPostFinal = interface(IDispatch)
   ['{CD000028-8B95-11D1-82DB-00C04FB1625D}']
    // OnPostFinal : Called by the NNTP event dispatcher after a posted message has been saved to its final location. 
   procedure OnPostFinal(Msg:IMessage;var EventStatus:CdoEventStatus);safecall;
  end;


// INNTPOnPostFinal : Implement when creating NNTP OnPostFinal event sinks.

 INNTPOnPostFinalDisp = dispinterface
   ['{CD000028-8B95-11D1-82DB-00C04FB1625D}']
    // OnPostFinal : Called by the NNTP event dispatcher after a posted message has been saved to its final location. 
   procedure OnPostFinal(Msg:IMessage;var EventStatus:CdoEventStatus);dispid 256;
  end;

//CoClasses
  CoMessage = Class
  Public
    Class Function Create: IMessage;
    Class Function CreateRemote(const MachineName: string): IMessage;
  end;

  CoConfiguration = Class
  Public
    Class Function Create: IConfiguration;
    Class Function CreateRemote(const MachineName: string): IConfiguration;
  end;

  CoDropDirectory = Class
  Public
    Class Function Create: IDropDirectory;
    Class Function CreateRemote(const MachineName: string): IDropDirectory;
  end;

  CoSMTPConnector = Class
  Public
    Class Function Create: ISMTPScriptConnector;
    Class Function CreateRemote(const MachineName: string): ISMTPScriptConnector;
  end;

  CoNNTPEarlyConnector = Class
  Public
    Class Function Create: INNTPEarlyScriptConnector;
    Class Function CreateRemote(const MachineName: string): INNTPEarlyScriptConnector;
  end;

  CoNNTPPostConnector = Class
  Public
    Class Function Create: INNTPPostScriptConnector;
    Class Function CreateRemote(const MachineName: string): INNTPPostScriptConnector;
  end;

  CoNNTPFinalConnector = Class
  Public
    Class Function Create: INNTPFinalScriptConnector;
    Class Function CreateRemote(const MachineName: string): INNTPFinalScriptConnector;
  end;

implementation

uses comobj;

Class Function CoMessage.Create: IMessage;
begin
  Result := CreateComObject(CLASS_Message) as IMessage;
end;

Class Function CoMessage.CreateRemote(const MachineName: string): IMessage;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Message) as IMessage;
end;

Class Function CoConfiguration.Create: IConfiguration;
begin
  Result := CreateComObject(CLASS_Configuration) as IConfiguration;
end;

Class Function CoConfiguration.CreateRemote(const MachineName: string): IConfiguration;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Configuration) as IConfiguration;
end;

Class Function CoDropDirectory.Create: IDropDirectory;
begin
  Result := CreateComObject(CLASS_DropDirectory) as IDropDirectory;
end;

Class Function CoDropDirectory.CreateRemote(const MachineName: string): IDropDirectory;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_DropDirectory) as IDropDirectory;
end;

Class Function CoSMTPConnector.Create: ISMTPScriptConnector;
begin
  Result := CreateComObject(CLASS_SMTPConnector) as ISMTPScriptConnector;
end;

Class Function CoSMTPConnector.CreateRemote(const MachineName: string): ISMTPScriptConnector;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_SMTPConnector) as ISMTPScriptConnector;
end;

Class Function CoNNTPEarlyConnector.Create: INNTPEarlyScriptConnector;
begin
  Result := CreateComObject(CLASS_NNTPEarlyConnector) as INNTPEarlyScriptConnector;
end;

Class Function CoNNTPEarlyConnector.CreateRemote(const MachineName: string): INNTPEarlyScriptConnector;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_NNTPEarlyConnector) as INNTPEarlyScriptConnector;
end;

Class Function CoNNTPPostConnector.Create: INNTPPostScriptConnector;
begin
  Result := CreateComObject(CLASS_NNTPPostConnector) as INNTPPostScriptConnector;
end;

Class Function CoNNTPPostConnector.CreateRemote(const MachineName: string): INNTPPostScriptConnector;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_NNTPPostConnector) as INNTPPostScriptConnector;
end;

Class Function CoNNTPFinalConnector.Create: INNTPFinalScriptConnector;
begin
  Result := CreateComObject(CLASS_NNTPFinalConnector) as INNTPFinalScriptConnector;
end;

Class Function CoNNTPFinalConnector.CreateRemote(const MachineName: string): INNTPFinalScriptConnector;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_NNTPFinalConnector) as INNTPFinalScriptConnector;
end;

end.
