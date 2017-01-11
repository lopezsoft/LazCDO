unit RegisterCDO;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,LResources, Forms, Controls, Graphics, Dialogs,
  CDOClass;

procedure Register;

implementation

procedure Register;
begin
  {$I cdomessage_icon.lrs}
  {$I cdoserver_icon.lrs}
  RegisterComponents('LazCDO',[TCDOMessage]);
  RegisterComponents('LazCDO',[TCDOServer]);
end;

end.

