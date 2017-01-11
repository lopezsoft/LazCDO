
{ TCDOServer }

procedure TCDOServer.SetSendPassword(AValue: String);
begin
  if FSendPassword=AValue then Exit;
  FSendPassword:=AValue;
end;

procedure TCDOServer.SetSchema(AValue: String);
begin
  if FSchema=AValue then Exit;
  FSchema:=AValue;
end;

procedure TCDOServer.SetActive(AValue: Boolean);
begin
  if FActive = AValue then Exit;
  FActive := AValue;
  if AValue then
    begin
     SetConfig;
     if SHandle = Nil then
       FActive:= False;
    end
  else
    SHandle:= Nil;
end;

procedure TCDOServer.SetSendUserName(AValue: String);
begin
  if FSendUserName=AValue then Exit;
  FSendUserName:=AValue;
end;

procedure TCDOServer.SetSendUsing(AValue: Integer);
begin
  if FSendUsing=AValue then Exit;
  FSendUsing:=AValue;
end;

procedure TCDOServer.SetSMTP(AValue: TSMTP);
begin
  if FSMTP=AValue then Exit;
  FSMTP := AValue;
  case FSMTP of
  sNone:
    begin
      SMTPServer      := 'smtp.mail.com';
      SMTPServerPort  := CDOPort;
      SMTPAuthenticate := False;
      SMTPUseSSL       := False;
    end;
  sGmail:
    begin
      SMTPServer       := 'smtp.gmail.com';
      SMTPServerPort   := CDOPortGmail;
      SMTPAuthenticate := True;
      SMTPUseSSL       := True;
    end;
  sHotmail:
    begin
      SMTPServer       := 'smtp.live.com';
      SMTPServerPort   := CDOPort;
      SMTPAuthenticate := True;
      SMTPUseSSL       := True;
    end;
  sOutlook:
    begin
      SMTPServer       := 'smtp-mail.outlook.com';
      SMTPServerPort   := CDOPortOutLook;
      SMTPAuthenticate := True;
      SMTPUseSSL       := False;
    end;
  sYahoo:
    begin
      SMTPServer       := 'smtp-mail.outlook.com';
      SMTPServerPort   := CDOPortYahho;
      SMTPAuthenticate := True;
      SMTPUseSSL       := False;
    end;
  end;
end;

procedure TCDOServer.SetSMTPAuthenticate(AValue: Boolean);
begin
  if FSMTPAuthenticate=AValue then Exit;
  FSMTPAuthenticate:=AValue;
end;

procedure TCDOServer.SetSMTPconnectionTimeOut(AValue: Integer);
begin
  if FSMTPconnectionTimeOut=AValue then Exit;
  FSMTPconnectionTimeOut:=AValue;
end;

procedure TCDOServer.SetSMTPServer(AValue: String);
begin
  if FSMTPServer=AValue then Exit;
  FSMTPServer:=AValue;
end;

procedure TCDOServer.SetSMTPServerPort(AValue: Integer);
begin
  if FSMTPServerPort=AValue then Exit;
  FSMTPServerPort:=AValue;
end;

procedure TCDOServer.SetSMTPUseSSL(AValue: Boolean);
begin
  if FSMTPUseSSL=AValue then Exit;
  FSMTPUseSSL:=AValue;
end;

function TCDOServer.GetDefaultSchema: String;
begin
  Result := CDOSchema;
end;

function TCDOServer.GetActive: Boolean;
begin
  Result := FActive;
end;

procedure TCDOServer.SetConfig;
var
  S  : String;
  M : IMessage;
  IC  : IConfiguration;
begin
  try
    if FSchema = '' then
      FSchema := CDOSchema;

    IC:= CoConfiguration.Create;

    with IC.FIELDS do begin
      ITEM[FSchema + 'smtpserver'].VALUE        := Self.SMTPServer;
      ITEM[FSchema + 'smtpserverport'].VALUE    := SMTPServerPort;
      ITEM[FSchema + 'sendusing'].VALUE         := SendUsing;
      ITEM[FSchema + 'smtpauthenticate'].VALUE  := SMTPAuthenticate;
      ITEM[FSchema + 'smtpusessl'].VALUE        := SMTPUseSSL;
      ITEM[FSchema + 'sendusername'].VALUE      := SendUserName;
      ITEM[FSchema + 'sendpassword'].VALUE      := SendPassword;
      ITEM[FSchema + 'smtpconnectiontimeout'].VALUE := SMTPconnectionTimeOut;
      UPDATE;
    end;

    SHandle := IC;
    IC := Nil;
  except on E : Exception do
    begin
      IC       := Nil;
      SHandle  := Nil;
      MessageDlg(CDO_MSG_NAME,CDO_MSG_ERROR_1+chr(13)+CDO_MSG_ERROR_2 + E.Message,mtError,[mbOK],0);
    end;
  end;
end;

constructor TCDOServer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  SMTPServer      := 'smtp.mail.com';
  SMTPServerPort  := CDOPort;
  SendPassword    := 'password';
  SendUserName    := 'user@mail.com';
  SendUsing       := 2;
  FSchema         := GetDefaultSchema;
  SHandle         := Nil;
  SMTPconnectionTimeOut := 30;
end;

destructor TCDOServer.Destroy;
begin
  Close;
  inherited Destroy;
end;

procedure TCDOServer.Assign(Source: TPersistent);
begin
  if (Source is TCDOServer) then
    begin
      FSendPassword   := TCDOServer(Source).FSendPassword;
      FSendUserName   := TCDOServer(Source).FSendPassword;
      FSendUsing      := TCDOServer(Source).FSendUsing;
      FSMTPAuthenticate := TCDOServer(Source).FSMTPAuthenticate;
      FSMTPServer       := TCDOServer(Source).FSMTPServer;
      FSMTPServerPort   := TCDOServer(Source).FSMTPServerPort;
      FSMTPUseSSL       := TCDOServer(Source).FSMTPUseSSL;
    end
  else
    inherited Assign(Source);
end;

procedure TCDOServer.Open;
begin
  Active:= True;
end;

procedure TCDOServer.Close;
begin
  Active:= False;
end;


{ TCDOMessage }

procedure TCDOMessage.SetServerConfig(AValue: TCDOServer);
begin
  if FServerConfig=AValue then Exit;
  FServerConfig:=AValue;
end;

procedure TCDOMessage.SetSubject(AValue: WideString);
begin
  if FSubject=AValue then Exit;
  FSubject:=AValue;
end;

procedure TCDOMessage.SetTextBody(AValue: WideString);
begin
  if FTextBody=AValue then Exit;
  FTextBody:=AValue;
end;

procedure TCDOMessage.SetTo_(AValue: WideString);
begin
  if FTo_=AValue then Exit;
  FTo_:=AValue;
end;

procedure TCDOMessage.CreateMHandle;
begin
  if MHandle = Nil then
    MHandle :=  CoMessage.Create;
end;

procedure TCDOMessage.ShowLoading;
begin
  FrmSplash := TFrmSplash.Create(Nil);
  FrmSplash.Show;
  FrmSplash.Update;
end;

procedure TCDOMessage.HideLoading;
begin
  FrmSplash.Close;
end;

procedure TCDOMessage.SetFrom(AValue: WideString);
begin
  if FFrom=AValue then Exit;
  FFrom:=AValue;
end;

procedure TCDOMessage.SetBCC(AValue: WideString);
begin
  if FBCC=AValue then Exit;
  FBCC:=AValue;
end;

procedure TCDOMessage.SetCC(AValue: WideString);
begin
  if FCC=AValue then Exit;
  FCC:=AValue;
end;

procedure TCDOMessage.SetHTMLBody(AValue: WideString);
begin
  if FHTMLBody=AValue then Exit;
  FHTMLBody:=AValue;
end;

constructor TCDOMessage.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  MHandle := Nil;
end;

destructor TCDOMessage.Destroy;
begin
  inherited Destroy;
end;

procedure TCDOMessage.Send;
var
  M : IMessage;
begin
  if Assigned(ServerConfig) then
    if ServerConfig.Active then
      begin
       try
         ShowLoading;
         CreateMHandle;
         M := MHandle;
         with M do
         begin
           Configuration := ServerConfig.SHandle;
           From   := Self.From;
           To_    := Self.To_;
           CC     := Self.CC;
           BCC    := Self.BCC;
           Subject:= Self.Subject;
           if Self.TextBody <> '' then
             TextBody := Self.TextBody;
           if Self.HTMLBody <> '' then
             HTMLBody := Self.HTMLBody;
           Send;
           HideLoading;
           if Assigned(FOnSend) then FOnSend(Self);
           MessageDlg(ApplicationName,CDO_MSG_SEND_OK,mtInformation,[mbOK],0);
         end;
       Except on E : Exception do
         begin
           HideLoading;
           MessageDlg(CDO_MSG_NAME,CDO_MSG_ERROR_1+CDO_MSG_ERROR_3+chr(13)+
                     CDO_MSG_ERROR_2+E.Message,mtError,[mbOK],0);
         end;
       end;
      end
  else
    MessageDlg(CDO_MSG_NAME,CDO_NOT_ASSIGN_SERVER,mtWarning,[mbOK],0);
end;

procedure TCDOMessage.AddAttachment(URL: WideString);
begin
  CreateMHandle;
  if Assigned(ServerConfig) then
    MHandle.AddAttachment(URL,ServerConfig.SendUserName,ServerConfig.SendPassword)
  else
    MessageDlg(CDO_MSG_NAME,CDO_NOT_ASSIGN_SERVER,mtWarning,[mbOK],0);
end;
