unit uFrmSplash;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  CDOCons;

type

  { TFrmSplash }

  TFrmSplash = class(TForm)
    lbMsg: TLabel;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FrmSplash: TFrmSplash;

implementation

{$R *.lfm}

{ TFrmSplash }

procedure TFrmSplash.FormCreate(Sender: TObject);
begin
   lbMsg.Caption:= CDO_MSG_SENDING;
end;

procedure TFrmSplash.FormActivate(Sender: TObject);
begin
  lbMsg.Caption:= CDO_MSG_SENDING;
end;

end.

