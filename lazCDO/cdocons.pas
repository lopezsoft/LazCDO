unit CDOCons;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources;

resourcestring
  CDO_MSG_LOAD_FILE = 'Abrir archivo...';
  CDO_MSG_SENDING = 'Enviando Email...';
  CDO_MSG_NAME = 'LazCDO Libreria';
  CDO_MSG_ERROR_1 = 'Ha ocurrido un error';
  CDO_MSG_ERROR_2 = 'Mensage de error: ';
  CDO_MSG_ERROR_3 = ' al tratar de enviar el correo';
  CDO_MSG_SEND_OK = 'El Email se ha enviado correctamente.';
  CDO_NOT_ASSIGN_SERVER = 'No está configurado el servidor.';

const
  CDOSchema = 'http://schemas.microsoft.com/cdo/configuration/';
  CDOPort   = 25;
  CDOPortGmail     = 465;
  CDOPortOutLook   = 587;
  CDOPortYahho     = 25;
implementation

end.

