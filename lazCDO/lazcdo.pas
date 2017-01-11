{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit lazCDO;

interface

uses
  CDOClass, RegisterCDO, CDOCons, uFrmSplash, LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('RegisterCDO', @RegisterCDO.Register);
end;

initialization
  RegisterPackage('lazCDO', @Register);
end.
