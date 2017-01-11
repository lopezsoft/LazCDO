unit CDOClass;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, CDOCons, comobj, variants, Dialogs,CDO_1_0_TLB,
  uFrmSplash;

type

  TSMTP = (sNone,sGmail,sHotmail,sOutlook,sYahoo);

  TCDOMessage = Class;
  TCDOServer  = Class;

  { TCDOServer }

  TCDOServer = class(TComponent)
  private
    FActive: Boolean;
    FSchema: String;
    FSendPassword: String;
    FSendUserName: String;
    FSendUsing: Integer;
    FSMTP: TSMTP;
    FSMTPAuthenticate: Boolean;
    FSMTPconnectionTimeOut: Integer;
    FSMTPServer: String;
    FSMTPServerPort: Integer;
    FSMTPUseSSL: Boolean;
    procedure SetActive(AValue: Boolean);
    procedure SetSchema(AValue: String);
    procedure SetSendPassword(AValue: String);
    procedure SetSendUserName(AValue: String);
    procedure SetSendUsing(AValue: Integer);
    procedure SetSMTP(AValue: TSMTP);
    procedure SetSMTPAuthenticate(AValue: Boolean);
    procedure SetSMTPconnectionTimeOut(AValue: Integer);
    procedure SetSMTPServer(AValue: String);
    procedure SetSMTPServerPort(AValue: Integer);
    procedure SetSMTPUseSSL(AValue: Boolean);
    function GetDefaultSchema : String;
    function GetActive : Boolean;
    { Private declarations }
  protected
    { Protected declarations }
    procedure SetConfig;
  public
    { Public declarations }
    SHandle : IConfiguration;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    procedure Open; virtual;
    procedure Close; virtual;
    property Schema : String read FSchema write SetSchema;
  published
    { Published declarations }
    property SMTPServer : String read FSMTPServer write SetSMTPServer;
    property SMTPServerPort : Integer read FSMTPServerPort write SetSMTPServerPort;
    property SendUsing : Integer read FSendUsing write SetSendUsing;
    property SMTPAuthenticate : Boolean read FSMTPAuthenticate write SetSMTPAuthenticate;
    property SMTPUseSSL : Boolean read FSMTPUseSSL write SetSMTPUseSSL;
    property SendUserName : String read FSendUserName write SetSendUserName;
    property SendPassword : String read FSendPassword write SetSendPassword;
    property SMTPconnectionTimeOut : Integer read FSMTPconnectionTimeOut
      write SetSMTPconnectionTimeOut;
    property Active : Boolean read GetActive write SetActive;
    property SMTP : TSMTP read FSMTP write SetSMTP;
  end;

  { TCDOMessage }

  TCDOMessage = class(TComponent)
  private
    FBCC: WideString;
    FCC: WideString;
    FFrom: WideString;
    FHTMLBody: WideString;
    FServerConfig: TCDOServer;
    FSubject: WideString;
    FTextBody: WideString;
    FTo_: WideString;
    FOnSend : TNotifyEvent;
    procedure SetBCC(AValue: WideString);
    procedure SetCC(AValue: WideString);
    procedure SetFrom(AValue: WideString);
    procedure SetHTMLBody(AValue: WideString);
    procedure SetServerConfig(AValue: TCDOServer);
    procedure SetSubject(AValue: WideString);
    procedure SetTextBody(AValue: WideString);
    procedure SetTo_(AValue: WideString);
    { Private declarations }
  protected
    { Protected declarations }
    procedure CreateMHandle;
    procedure ShowLoading;
    procedure HideLoading;
  public
    { Public declarations }
    MHandle : IMessage;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Send;virtual;
    procedure AddAttachment(URL : WideString);
  published
    { Published declarations }
    property ServerConfig : TCDOServer read FServerConfig write FServerConfig;
    property From : WideString read FFrom write SetFrom;
    property To_  : WideString read FTo_ write SetTo_;
    property CC   : WideString read FCC write SetCC;
    property BCC  : WideString read FBCC write SetBCC;
    property Subject : WideString read FSubject write SetSubject;
    property TextBody : WideString read FTextBody write SetTextBody;
    property HTMLBody : WideString read FHTMLBody write SetHTMLBody;
    property OnSend   : TNotifyEvent read FOnSend write FOnSend;
  end;

implementation

{$I CDOClass.inc}

end.

