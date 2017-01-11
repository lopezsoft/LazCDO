Unit ADODB_6_0_TLB;

//  Imported ADODB on 23/12/2016 2:21:37 from C:\Program Files (x86)\Common Files\System\ado\msado60.tlb

{$mode delphi}{$H+}

interface

//  Warning: renamed interface 'Property' to 'Property_'
//  Warning: renamed coclass 'Record' to 'Record_'
//  Warning: renamed parameter 'Object' in _DynaCollection.Append to 'Object_'
//  Warning: renamed parameter 'Object' in _DynaCollection.Append to 'Object_'
//  Warning: renamed property 'Type' in Property_ to 'Type_'
//  Warning: renamed parameter 'Type' in Command15.CreateParameter to 'Type_'
//  Warning: renamed parameter 'Type' in Command15.CreateParameter to 'Type_'
//  Warning: renamed property 'EOF' in _Recordset to 'EOF_'
//  Warning: renamed property 'EOF' in Recordset21 to 'EOF_'
//  Warning: renamed property 'EOF' in Recordset20 to 'EOF_'
//  Warning: renamed property 'EOF' in Recordset15 to 'EOF_'
//  Warning: renamed parameter 'Type' in Fields.Append to 'Type_'
//  Warning: renamed parameter 'Type' in Fields._Append to 'Type_'
//  Warning: renamed parameter 'Type' in Fields.Append to 'Type_'
//  Warning: renamed parameter 'Type' in Fields20._Append to 'Type_'
//  Warning: renamed parameter 'Type' in Fields20._Append to 'Type_'
//  Warning: renamed property 'Type' in Field to 'Type_'
//  Warning: renamed property 'Type' in Field20 to 'Type_'
//  Warning: renamed property 'Type' in _Parameter to 'Type_'
//  Warning: renamed parameter 'Object' in Parameters.Append to 'Object_'
//  Warning: renamed parameter 'Type' in Command25.CreateParameter to 'Type_'
//  Warning: renamed parameter 'Type' in _Command.CreateParameter to 'Type_'
//  Warning: renamed property 'Type' in _Stream to 'Type_'
//  Warning: renamed method 'Read' in _Stream to 'Read_'
//  Warning: renamed method 'Write' in _Stream to 'Write_'
//  Warning: renamed method 'Read' in _Stream to 'Read_'
//  Warning: renamed method 'Write' in _Stream to 'Write_'
//  Warning: renamed property 'Type' in Field15 to 'Type_'
Uses
  Windows,ActiveX,Classes,Variants,EventSink;
Const
  ADODBMajorVersion = 6;
  ADODBMinorVersion = 0;
  ADODBLCID = 0;
  LIBID_ADODB : TGUID = '{B691E011-1797-432E-907A-4D8C69339129}';

  IID__Collection : TGUID = '{00000512-0000-0010-8000-00AA006D2EA4}';
  IID__DynaCollection : TGUID = '{00000513-0000-0010-8000-00AA006D2EA4}';
  IID__ADO : TGUID = '{00000534-0000-0010-8000-00AA006D2EA4}';
  IID_Properties : TGUID = '{00000504-0000-0010-8000-00AA006D2EA4}';
  IID_Property : TGUID = '{00000503-0000-0010-8000-00AA006D2EA4}';
  IID_Error : TGUID = '{00000500-0000-0010-8000-00AA006D2EA4}';
  IID_Errors : TGUID = '{00000501-0000-0010-8000-00AA006D2EA4}';
  IID_Command15 : TGUID = '{00000508-0000-0010-8000-00AA006D2EA4}';
  IID__Connection : TGUID = '{00000550-0000-0010-8000-00AA006D2EA4}';
  IID_Connection15 : TGUID = '{00000515-0000-0010-8000-00AA006D2EA4}';
  IID__Recordset : TGUID = '{00000556-0000-0010-8000-00AA006D2EA4}';
  IID_Recordset21 : TGUID = '{00000555-0000-0010-8000-00AA006D2EA4}';
  IID_Recordset20 : TGUID = '{0000054F-0000-0010-8000-00AA006D2EA4}';
  IID_Recordset15 : TGUID = '{0000050E-0000-0010-8000-00AA006D2EA4}';
  IID_Fields : TGUID = '{00000564-0000-0010-8000-00AA006D2EA4}';
  IID_Fields20 : TGUID = '{0000054D-0000-0010-8000-00AA006D2EA4}';
  IID_Fields15 : TGUID = '{00000506-0000-0010-8000-00AA006D2EA4}';
  IID_Field : TGUID = '{00000569-0000-0010-8000-00AA006D2EA4}';
  IID_Field20 : TGUID = '{0000054C-0000-0010-8000-00AA006D2EA4}';
  IID__Parameter : TGUID = '{0000050C-0000-0010-8000-00AA006D2EA4}';
  IID_Parameters : TGUID = '{0000050D-0000-0010-8000-00AA006D2EA4}';
  IID_Command25 : TGUID = '{0000054E-0000-0010-8000-00AA006D2EA4}';
  IID__Command : TGUID = '{B08400BD-F9D1-4D02-B856-71D5DBA123E9}';
  IID_ConnectionEventsVt : TGUID = '{00000402-0000-0010-8000-00AA006D2EA4}';
  IID_RecordsetEventsVt : TGUID = '{00000403-0000-0010-8000-00AA006D2EA4}';
  IID_ConnectionEvents : TGUID = '{00000400-0000-0010-8000-00AA006D2EA4}';
  IID_RecordsetEvents : TGUID = '{00000266-0000-0010-8000-00AA006D2EA4}';
  IID_ADOConnectionConstruction15 : TGUID = '{00000516-0000-0010-8000-00AA006D2EA4}';
  IID_ADOConnectionConstruction : TGUID = '{00000551-0000-0010-8000-00AA006D2EA4}';
  CLASS_Connection : TGUID = '{00000514-0000-0010-8000-00AA006D2EA4}';
  IID__Record : TGUID = '{00000562-0000-0010-8000-00AA006D2EA4}';
  CLASS_Record : TGUID = '{00000560-0000-0010-8000-00AA006D2EA4}';
  IID__Stream : TGUID = '{00000565-0000-0010-8000-00AA006D2EA4}';
  CLASS_Stream : TGUID = '{00000566-0000-0010-8000-00AA006D2EA4}';
  IID_ADORecordConstruction : TGUID = '{00000567-0000-0010-8000-00AA006D2EA4}';
  IID_ADOStreamConstruction : TGUID = '{00000568-0000-0010-8000-00AA006D2EA4}';
  IID_ADOCommandConstruction : TGUID = '{00000517-0000-0010-8000-00AA006D2EA4}';
  CLASS_Command : TGUID = '{00000507-0000-0010-8000-00AA006D2EA4}';
  CLASS_Recordset : TGUID = '{00000535-0000-0010-8000-00AA006D2EA4}';
  IID_ADORecordsetConstruction : TGUID = '{00000283-0000-0010-8000-00AA006D2EA4}';
  IID_Field15 : TGUID = '{00000505-0000-0010-8000-00AA006D2EA4}';
  CLASS_Parameter : TGUID = '{0000050B-0000-0010-8000-00AA006D2EA4}';

//Enums

Type
  CursorTypeEnum =LongWord;
Const
  adOpenUnspecified = $FFFFFFFFFFFFFFFF;
  adOpenForwardOnly = $0000000000000000;
  adOpenKeyset = $0000000000000001;
  adOpenDynamic = $0000000000000002;
  adOpenStatic = $0000000000000003;
Type
  CursorOptionEnum =LongWord;
Const
  adHoldRecords = $0000000000000100;
  adMovePrevious = $0000000000000200;
  adAddNew = $0000000001000400;
  adDelete = $0000000001000800;
  adUpdate = $0000000001008000;
  adBookmark = $0000000000002000;
  adApproxPosition = $0000000000004000;
  adUpdateBatch = $0000000000010000;
  adResync = $0000000000020000;
  adNotify = $0000000000040000;
  adFind = $0000000000080000;
  adSeek = $0000000000400000;
  adIndex = $0000000000800000;
Type
  LockTypeEnum =LongWord;
Const
  adLockUnspecified = $FFFFFFFFFFFFFFFF;
  adLockReadOnly = $0000000000000001;
  adLockPessimistic = $0000000000000002;
  adLockOptimistic = $0000000000000003;
  adLockBatchOptimistic = $0000000000000004;
Type
  ExecuteOptionEnum =LongWord;
Const
  adOptionUnspecified = $FFFFFFFFFFFFFFFF;
  adAsyncExecute = $0000000000000010;
  adAsyncFetch = $0000000000000020;
  adAsyncFetchNonBlocking = $0000000000000040;
  adExecuteNoRecords = $0000000000000080;
  adExecuteStream = $0000000000000400;
  adExecuteRecord = $0000000000000800;
Type
  ConnectOptionEnum =LongWord;
Const
  adConnectUnspecified = $FFFFFFFFFFFFFFFF;
  adAsyncConnect = $0000000000000010;
Type
  ObjectStateEnum =LongWord;
Const
  adStateClosed = $0000000000000000;
  adStateOpen = $0000000000000001;
  adStateConnecting = $0000000000000002;
  adStateExecuting = $0000000000000004;
  adStateFetching = $0000000000000008;
Type
  CursorLocationEnum =LongWord;
Const
  adUseNone = $0000000000000001;
  adUseServer = $0000000000000002;
  adUseClient = $0000000000000003;
  adUseClientBatch = $0000000000000003;
Type
  DataTypeEnum =LongWord;
Const
  adEmpty = $0000000000000000;
  adTinyInt = $0000000000000010;
  adSmallInt = $0000000000000002;
  adInteger = $0000000000000003;
  adBigInt = $0000000000000014;
  adUnsignedTinyInt = $0000000000000011;
  adUnsignedSmallInt = $0000000000000012;
  adUnsignedInt = $0000000000000013;
  adUnsignedBigInt = $0000000000000015;
  adSingle = $0000000000000004;
  adDouble = $0000000000000005;
  adCurrency = $0000000000000006;
  adDecimal = $000000000000000E;
  adNumeric = $0000000000000083;
  adBoolean = $000000000000000B;
  adError = $000000000000000A;
  adUserDefined = $0000000000000084;
  adVariant = $000000000000000C;
  adIDispatch = $0000000000000009;
  adIUnknown = $000000000000000D;
  adGUID = $0000000000000048;
  adDate = $0000000000000007;
  adDBDate = $0000000000000085;
  adDBTime = $0000000000000086;
  adDBTimeStamp = $0000000000000087;
  adBSTR = $0000000000000008;
  adChar = $0000000000000081;
  adVarChar = $00000000000000C8;
  adLongVarChar = $00000000000000C9;
  adWChar = $0000000000000082;
  adVarWChar = $00000000000000CA;
  adLongVarWChar = $00000000000000CB;
  adBinary = $0000000000000080;
  adVarBinary = $00000000000000CC;
  adLongVarBinary = $00000000000000CD;
  adChapter = $0000000000000088;
  adFileTime = $0000000000000040;
  adPropVariant = $000000000000008A;
  adVarNumeric = $000000000000008B;
  adArray = $0000000000002000;
Type
  FieldAttributeEnum =LongWord;
Const
  adFldUnspecified = $FFFFFFFFFFFFFFFF;
  adFldMayDefer = $0000000000000002;
  adFldUpdatable = $0000000000000004;
  adFldUnknownUpdatable = $0000000000000008;
  adFldFixed = $0000000000000010;
  adFldIsNullable = $0000000000000020;
  adFldMayBeNull = $0000000000000040;
  adFldLong = $0000000000000080;
  adFldRowID = $0000000000000100;
  adFldRowVersion = $0000000000000200;
  adFldCacheDeferred = $0000000000001000;
  adFldIsChapter = $0000000000002000;
  adFldNegativeScale = $0000000000004000;
  adFldKeyColumn = $0000000000008000;
  adFldIsRowURL = $0000000000010000;
  adFldIsDefaultStream = $0000000000020000;
  adFldIsCollection = $0000000000040000;
Type
  EditModeEnum =LongWord;
Const
  adEditNone = $0000000000000000;
  adEditInProgress = $0000000000000001;
  adEditAdd = $0000000000000002;
  adEditDelete = $0000000000000004;
Type
  RecordStatusEnum =LongWord;
Const
  adRecOK = $0000000000000000;
  adRecNew = $0000000000000001;
  adRecModified = $0000000000000002;
  adRecDeleted = $0000000000000004;
  adRecUnmodified = $0000000000000008;
  adRecInvalid = $0000000000000010;
  adRecMultipleChanges = $0000000000000040;
  adRecPendingChanges = $0000000000000080;
  adRecCanceled = $0000000000000100;
  adRecCantRelease = $0000000000000400;
  adRecConcurrencyViolation = $0000000000000800;
  adRecIntegrityViolation = $0000000000001000;
  adRecMaxChangesExceeded = $0000000000002000;
  adRecObjectOpen = $0000000000004000;
  adRecOutOfMemory = $0000000000008000;
  adRecPermissionDenied = $0000000000010000;
  adRecSchemaViolation = $0000000000020000;
  adRecDBDeleted = $0000000000040000;
Type
  GetRowsOptionEnum =LongWord;
Const
  adGetRowsRest = $FFFFFFFFFFFFFFFF;
Type
  PositionEnum =LongWord;
Const
  adPosUnknown = $FFFFFFFFFFFFFFFF;
  adPosBOF = $FFFFFFFFFFFFFFFE;
  adPosEOF = $FFFFFFFFFFFFFFFD;
Type
  BookmarkEnum =LongWord;
Const
  adBookmarkCurrent = $0000000000000000;
  adBookmarkFirst = $0000000000000001;
  adBookmarkLast = $0000000000000002;
Type
  MarshalOptionsEnum =LongWord;
Const
  adMarshalAll = $0000000000000000;
  adMarshalModifiedOnly = $0000000000000001;
Type
  AffectEnum =LongWord;
Const
  adAffectCurrent = $0000000000000001;
  adAffectGroup = $0000000000000002;
  adAffectAll = $0000000000000003;
  adAffectAllChapters = $0000000000000004;
Type
  ResyncEnum =LongWord;
Const
  adResyncUnderlyingValues = $0000000000000001;
  adResyncAllValues = $0000000000000002;
Type
  CompareEnum =LongWord;
Const
  adCompareLessThan = $0000000000000000;
  adCompareEqual = $0000000000000001;
  adCompareGreaterThan = $0000000000000002;
  adCompareNotEqual = $0000000000000003;
  adCompareNotComparable = $0000000000000004;
Type
  FilterGroupEnum =LongWord;
Const
  adFilterNone = $0000000000000000;
  adFilterPendingRecords = $0000000000000001;
  adFilterAffectedRecords = $0000000000000002;
  adFilterFetchedRecords = $0000000000000003;
  adFilterPredicate = $0000000000000004;
  adFilterConflictingRecords = $0000000000000005;
Type
  SearchDirectionEnum =LongWord;
Const
  adSearchForward = $0000000000000001;
  adSearchBackward = $FFFFFFFFFFFFFFFF;
Type
  PersistFormatEnum =LongWord;
Const
  adPersistADTG = $0000000000000000;
  adPersistXML = $0000000000000001;
Type
  StringFormatEnum =LongWord;
Const
  adClipString = $0000000000000002;
Type
  ConnectPromptEnum =LongWord;
Const
  adPromptAlways = $0000000000000001;
  adPromptComplete = $0000000000000002;
  adPromptCompleteRequired = $0000000000000003;
  adPromptNever = $0000000000000004;
Type
  ConnectModeEnum =LongWord;
Const
  adModeUnknown = $0000000000000000;
  adModeRead = $0000000000000001;
  adModeWrite = $0000000000000002;
  adModeReadWrite = $0000000000000003;
  adModeShareDenyRead = $0000000000000004;
  adModeShareDenyWrite = $0000000000000008;
  adModeShareExclusive = $000000000000000C;
  adModeShareDenyNone = $0000000000000010;
  adModeRecursive = $0000000000400000;
Type
  RecordCreateOptionsEnum =LongWord;
Const
  adCreateCollection = $0000000000002000;
  adCreateStructDoc = $FFFFFFFF80000000;
  adCreateNonCollection = $0000000000000000;
  adOpenIfExists = $0000000002000000;
  adCreateOverwrite = $0000000004000000;
  adFailIfNotExists = $FFFFFFFFFFFFFFFF;
Type
  RecordOpenOptionsEnum =LongWord;
Const
  adOpenRecordUnspecified = $FFFFFFFFFFFFFFFF;
  adOpenSource = $0000000000800000;
  adOpenOutput = $0000000000800000;
  adOpenAsync = $0000000000001000;
  adDelayFetchStream = $0000000000004000;
  adDelayFetchFields = $0000000000008000;
  adOpenExecuteCommand = $0000000000010000;
Type
  IsolationLevelEnum =LongWord;
Const
  adXactUnspecified = $FFFFFFFFFFFFFFFF;
  adXactChaos = $0000000000000010;
  adXactReadUncommitted = $0000000000000100;
  adXactBrowse = $0000000000000100;
  adXactCursorStability = $0000000000001000;
  adXactReadCommitted = $0000000000001000;
  adXactRepeatableRead = $0000000000010000;
  adXactSerializable = $0000000000100000;
  adXactIsolated = $0000000000100000;
Type
  XactAttributeEnum =LongWord;
Const
  adXactCommitRetaining = $0000000000020000;
  adXactAbortRetaining = $0000000000040000;
  adXactAsyncPhaseOne = $0000000000080000;
  adXactSyncPhaseOne = $0000000000100000;
Type
  PropertyAttributesEnum =LongWord;
Const
  adPropNotSupported = $0000000000000000;
  adPropRequired = $0000000000000001;
  adPropOptional = $0000000000000002;
  adPropRead = $0000000000000200;
  adPropWrite = $0000000000000400;
Type
  ErrorValueEnum =LongWord;
Const
  adErrProviderFailed = $0000000000000BB8;
  adErrInvalidArgument = $0000000000000BB9;
  adErrOpeningFile = $0000000000000BBA;
  adErrReadFile = $0000000000000BBB;
  adErrWriteFile = $0000000000000BBC;
  adErrNoCurrentRecord = $0000000000000BCD;
  adErrIllegalOperation = $0000000000000C93;
  adErrCantChangeProvider = $0000000000000C94;
  adErrInTransaction = $0000000000000CAE;
  adErrFeatureNotAvailable = $0000000000000CB3;
  adErrItemNotFound = $0000000000000CC1;
  adErrObjectInCollection = $0000000000000D27;
  adErrObjectNotSet = $0000000000000D5C;
  adErrDataConversion = $0000000000000D5D;
  adErrObjectClosed = $0000000000000E78;
  adErrObjectOpen = $0000000000000E79;
  adErrProviderNotFound = $0000000000000E7A;
  adErrBoundToCommand = $0000000000000E7B;
  adErrInvalidParamInfo = $0000000000000E7C;
  adErrInvalidConnection = $0000000000000E7D;
  adErrNotReentrant = $0000000000000E7E;
  adErrStillExecuting = $0000000000000E7F;
  adErrOperationCancelled = $0000000000000E80;
  adErrStillConnecting = $0000000000000E81;
  adErrInvalidTransaction = $0000000000000E82;
  adErrNotExecuting = $0000000000000E83;
  adErrUnsafeOperation = $0000000000000E84;
  adwrnSecurityDialog = $0000000000000E85;
  adwrnSecurityDialogHeader = $0000000000000E86;
  adErrIntegrityViolation = $0000000000000E87;
  adErrPermissionDenied = $0000000000000E88;
  adErrDataOverflow = $0000000000000E89;
  adErrSchemaViolation = $0000000000000E8A;
  adErrSignMismatch = $0000000000000E8B;
  adErrCantConvertvalue = $0000000000000E8C;
  adErrCantCreate = $0000000000000E8D;
  adErrColumnNotOnThisRow = $0000000000000E8E;
  adErrURLDoesNotExist = $0000000000000E8F;
  adErrTreePermissionDenied = $0000000000000E90;
  adErrInvalidURL = $0000000000000E91;
  adErrResourceLocked = $0000000000000E92;
  adErrResourceExists = $0000000000000E93;
  adErrCannotComplete = $0000000000000E94;
  adErrVolumeNotFound = $0000000000000E95;
  adErrOutOfSpace = $0000000000000E96;
  adErrResourceOutOfScope = $0000000000000E97;
  adErrUnavailable = $0000000000000E98;
  adErrURLNamedRowDoesNotExist = $0000000000000E99;
  adErrDelResOutOfScope = $0000000000000E9A;
  adErrPropInvalidColumn = $0000000000000E9B;
  adErrPropInvalidOption = $0000000000000E9C;
  adErrPropInvalidValue = $0000000000000E9D;
  adErrPropConflicting = $0000000000000E9E;
  adErrPropNotAllSettable = $0000000000000E9F;
  adErrPropNotSet = $0000000000000EA0;
  adErrPropNotSettable = $0000000000000EA1;
  adErrPropNotSupported = $0000000000000EA2;
  adErrCatalogNotSet = $0000000000000EA3;
  adErrCantChangeConnection = $0000000000000EA4;
  adErrFieldsUpdateFailed = $0000000000000EA5;
  adErrDenyNotSupported = $0000000000000EA6;
  adErrDenyTypeNotSupported = $0000000000000EA7;
  adErrProviderNotSpecified = $0000000000000EA9;
  adErrConnectionStringTooLong = $0000000000000EAA;
Type
  ParameterAttributesEnum =LongWord;
Const
  adParamSigned = $0000000000000010;
  adParamNullable = $0000000000000040;
  adParamLong = $0000000000000080;
Type
  ParameterDirectionEnum =LongWord;
Const
  adParamUnknown = $0000000000000000;
  adParamInput = $0000000000000001;
  adParamOutput = $0000000000000002;
  adParamInputOutput = $0000000000000003;
  adParamReturnValue = $0000000000000004;
Type
  CommandTypeEnum =LongWord;
Const
  adCmdUnspecified = $FFFFFFFFFFFFFFFF;
  adCmdUnknown = $0000000000000008;
  adCmdText = $0000000000000001;
  adCmdTable = $0000000000000002;
  adCmdStoredProc = $0000000000000004;
  adCmdFile = $0000000000000100;
  adCmdTableDirect = $0000000000000200;
Type
  EventStatusEnum =LongWord;
Const
  adStatusOK = $0000000000000001;
  adStatusErrorsOccurred = $0000000000000002;
  adStatusCantDeny = $0000000000000003;
  adStatusCancel = $0000000000000004;
  adStatusUnwantedEvent = $0000000000000005;
Type
  EventReasonEnum =LongWord;
Const
  adRsnAddNew = $0000000000000001;
  adRsnDelete = $0000000000000002;
  adRsnUpdate = $0000000000000003;
  adRsnUndoUpdate = $0000000000000004;
  adRsnUndoAddNew = $0000000000000005;
  adRsnUndoDelete = $0000000000000006;
  adRsnRequery = $0000000000000007;
  adRsnResynch = $0000000000000008;
  adRsnClose = $0000000000000009;
  adRsnMove = $000000000000000A;
  adRsnFirstChange = $000000000000000B;
  adRsnMoveFirst = $000000000000000C;
  adRsnMoveNext = $000000000000000D;
  adRsnMovePrevious = $000000000000000E;
  adRsnMoveLast = $000000000000000F;
Type
  SchemaEnum =LongWord;
Const
  adSchemaProviderSpecific = $FFFFFFFFFFFFFFFF;
  adSchemaAsserts = $0000000000000000;
  adSchemaCatalogs = $0000000000000001;
  adSchemaCharacterSets = $0000000000000002;
  adSchemaCollations = $0000000000000003;
  adSchemaColumns = $0000000000000004;
  adSchemaCheckConstraints = $0000000000000005;
  adSchemaConstraintColumnUsage = $0000000000000006;
  adSchemaConstraintTableUsage = $0000000000000007;
  adSchemaKeyColumnUsage = $0000000000000008;
  adSchemaReferentialContraints = $0000000000000009;
  adSchemaReferentialConstraints = $0000000000000009;
  adSchemaTableConstraints = $000000000000000A;
  adSchemaColumnsDomainUsage = $000000000000000B;
  adSchemaIndexes = $000000000000000C;
  adSchemaColumnPrivileges = $000000000000000D;
  adSchemaTablePrivileges = $000000000000000E;
  adSchemaUsagePrivileges = $000000000000000F;
  adSchemaProcedures = $0000000000000010;
  adSchemaSchemata = $0000000000000011;
  adSchemaSQLLanguages = $0000000000000012;
  adSchemaStatistics = $0000000000000013;
  adSchemaTables = $0000000000000014;
  adSchemaTranslations = $0000000000000015;
  adSchemaProviderTypes = $0000000000000016;
  adSchemaViews = $0000000000000017;
  adSchemaViewColumnUsage = $0000000000000018;
  adSchemaViewTableUsage = $0000000000000019;
  adSchemaProcedureParameters = $000000000000001A;
  adSchemaForeignKeys = $000000000000001B;
  adSchemaPrimaryKeys = $000000000000001C;
  adSchemaProcedureColumns = $000000000000001D;
  adSchemaDBInfoKeywords = $000000000000001E;
  adSchemaDBInfoLiterals = $000000000000001F;
  adSchemaCubes = $0000000000000020;
  adSchemaDimensions = $0000000000000021;
  adSchemaHierarchies = $0000000000000022;
  adSchemaLevels = $0000000000000023;
  adSchemaMeasures = $0000000000000024;
  adSchemaProperties = $0000000000000025;
  adSchemaMembers = $0000000000000026;
  adSchemaTrustees = $0000000000000027;
  adSchemaFunctions = $0000000000000028;
  adSchemaActions = $0000000000000029;
  adSchemaCommands = $000000000000002A;
  adSchemaSets = $000000000000002B;
Type
  FieldStatusEnum =LongWord;
Const
  adFieldOK = $0000000000000000;
  adFieldCantConvertValue = $0000000000000002;
  adFieldIsNull = $0000000000000003;
  adFieldTruncated = $0000000000000004;
  adFieldSignMismatch = $0000000000000005;
  adFieldDataOverflow = $0000000000000006;
  adFieldCantCreate = $0000000000000007;
  adFieldUnavailable = $0000000000000008;
  adFieldPermissionDenied = $0000000000000009;
  adFieldIntegrityViolation = $000000000000000A;
  adFieldSchemaViolation = $000000000000000B;
  adFieldBadStatus = $000000000000000C;
  adFieldDefault = $000000000000000D;
  adFieldIgnore = $000000000000000F;
  adFieldDoesNotExist = $0000000000000010;
  adFieldInvalidURL = $0000000000000011;
  adFieldResourceLocked = $0000000000000012;
  adFieldResourceExists = $0000000000000013;
  adFieldCannotComplete = $0000000000000014;
  adFieldVolumeNotFound = $0000000000000015;
  adFieldOutOfSpace = $0000000000000016;
  adFieldCannotDeleteSource = $0000000000000017;
  adFieldReadOnly = $0000000000000018;
  adFieldResourceOutOfScope = $0000000000000019;
  adFieldAlreadyExists = $000000000000001A;
  adFieldPendingInsert = $0000000000010000;
  adFieldPendingDelete = $0000000000020000;
  adFieldPendingChange = $0000000000040000;
  adFieldPendingUnknown = $0000000000080000;
  adFieldPendingUnknownDelete = $0000000000100000;
Type
  SeekEnum =LongWord;
Const
  adSeekFirstEQ = $0000000000000001;
  adSeekLastEQ = $0000000000000002;
  adSeekAfterEQ = $0000000000000004;
  adSeekAfter = $0000000000000008;
  adSeekBeforeEQ = $0000000000000010;
  adSeekBefore = $0000000000000020;
Type
  ADCPROP_UPDATECRITERIA_ENUM =LongWord;
Const
  adCriteriaKey = $0000000000000000;
  adCriteriaAllCols = $0000000000000001;
  adCriteriaUpdCols = $0000000000000002;
  adCriteriaTimeStamp = $0000000000000003;
Type
  ADCPROP_ASYNCTHREADPRIORITY_ENUM =LongWord;
Const
  adPriorityLowest = $0000000000000001;
  adPriorityBelowNormal = $0000000000000002;
  adPriorityNormal = $0000000000000003;
  adPriorityAboveNormal = $0000000000000004;
  adPriorityHighest = $0000000000000005;
Type
  ADCPROP_AUTORECALC_ENUM =LongWord;
Const
  adRecalcUpFront = $0000000000000000;
  adRecalcAlways = $0000000000000001;
Type
  ADCPROP_UPDATERESYNC_ENUM =LongWord;
Const
  adResyncNone = $0000000000000000;
  adResyncAutoIncrement = $0000000000000001;
  adResyncConflicts = $0000000000000002;
  adResyncUpdates = $0000000000000004;
  adResyncInserts = $0000000000000008;
  adResyncAll = $000000000000000F;
Type
  MoveRecordOptionsEnum =LongWord;
Const
  adMoveUnspecified = $FFFFFFFFFFFFFFFF;
  adMoveOverWrite = $0000000000000001;
  adMoveDontUpdateLinks = $0000000000000002;
  adMoveAllowEmulation = $0000000000000004;
Type
  CopyRecordOptionsEnum =LongWord;
Const
  adCopyUnspecified = $FFFFFFFFFFFFFFFF;
  adCopyOverWrite = $0000000000000001;
  adCopyAllowEmulation = $0000000000000004;
  adCopyNonRecursive = $0000000000000002;
Type
  StreamTypeEnum =LongWord;
Const
  adTypeBinary = $0000000000000001;
  adTypeText = $0000000000000002;
Type
  LineSeparatorEnum =LongWord;
Const
  adLF = $000000000000000A;
  adCR = $000000000000000D;
  adCRLF = $FFFFFFFFFFFFFFFF;
Type
  StreamOpenOptionsEnum =LongWord;
Const
  adOpenStreamUnspecified = $FFFFFFFFFFFFFFFF;
  adOpenStreamAsync = $0000000000000001;
  adOpenStreamFromRecord = $0000000000000004;
Type
  StreamWriteEnum =LongWord;
Const
  adWriteChar = $0000000000000000;
  adWriteLine = $0000000000000001;
  stWriteChar = $0000000000000000;
  stWriteLine = $0000000000000001;
Type
  SaveOptionsEnum =LongWord;
Const
  adSaveCreateNotExist = $0000000000000001;
  adSaveCreateOverWrite = $0000000000000002;
Type
  FieldEnum =LongWord;
Const
  adDefaultStream = $FFFFFFFFFFFFFFFF;
  adRecordURL = $FFFFFFFFFFFFFFFE;
Type
  StreamReadEnum =LongWord;
Const
  adReadAll = $FFFFFFFFFFFFFFFF;
  adReadLine = $FFFFFFFFFFFFFFFE;
Type
  RecordTypeEnum =LongWord;
Const
  adSimpleRecord = $0000000000000000;
  adCollectionRecord = $0000000000000001;
  adStructDoc = $0000000000000002;
//Forward declarations

Type
 _Collection = interface;
 _CollectionDisp = dispinterface;
 _DynaCollection = interface;
 _DynaCollectionDisp = dispinterface;
 _ADO = interface;
 _ADODisp = dispinterface;
 Properties = interface;
 PropertiesDisp = dispinterface;
 Property_ = interface;
 Property_Disp = dispinterface;
 Error = interface;
 ErrorDisp = dispinterface;
 Errors = interface;
 ErrorsDisp = dispinterface;
 Command15 = interface;
 Command15Disp = dispinterface;
 _Connection = interface;
 _ConnectionDisp = dispinterface;
 Connection15 = interface;
 Connection15Disp = dispinterface;
 _Recordset = interface;
 _RecordsetDisp = dispinterface;
 Recordset21 = interface;
 Recordset21Disp = dispinterface;
 Recordset20 = interface;
 Recordset20Disp = dispinterface;
 Recordset15 = interface;
 Recordset15Disp = dispinterface;
 Fields = interface;
 FieldsDisp = dispinterface;
 Fields20 = interface;
 Fields20Disp = dispinterface;
 Fields15 = interface;
 Fields15Disp = dispinterface;
 Field = interface;
 FieldDisp = dispinterface;
 Field20 = interface;
 Field20Disp = dispinterface;
 _Parameter = interface;
 _ParameterDisp = dispinterface;
 Parameters = interface;
 ParametersDisp = dispinterface;
 Command25 = interface;
 Command25Disp = dispinterface;
 _Command = interface;
 _CommandDisp = dispinterface;
 ConnectionEventsVt = interface;
 RecordsetEventsVt = interface;
 ConnectionEvents = dispinterface;
 RecordsetEvents = dispinterface;
 ADOConnectionConstruction15 = interface;
 ADOConnectionConstruction = interface;
 _Record = interface;
 _RecordDisp = dispinterface;
 _Stream = interface;
 _StreamDisp = dispinterface;
 ADORecordConstruction = interface;
 ADOStreamConstruction = interface;
 ADOCommandConstruction = interface;
 ADORecordsetConstruction = interface;
 Field15 = interface;
 Field15Disp = dispinterface;

//Map CoClass to its default interface

 Connection = _Connection;
 Record_ = _Record;
 Stream = _Stream;
 Command = _Command;
 Recordset = _Recordset;
 Parameter = _Parameter;

//records, unions, aliases

 ADO_LONGPTR = Integer;
 PositionEnum_Param = PositionEnum;
 SearchDirection = SearchDirectionEnum;

//interface declarations

// _Collection : 

 _Collection = interface(IDispatch)
   ['{00000512-0000-0010-8000-00AA006D2EA4}']
   function Get_Count : Integer; safecall;
    // _NewEnum :  
   function _NewEnum:IUnknown;safecall;
    // Refresh :  
   procedure Refresh;safecall;
    // Count :  
   property Count:Integer read Get_Count;
  end;


// _Collection : 

 _CollectionDisp = dispinterface
   ['{00000512-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Count :  
   property Count:Integer  readonly dispid 1;
  end;


// _DynaCollection : 

 _DynaCollection = interface(_Collection)
   ['{00000513-0000-0010-8000-00AA006D2EA4}']
    // Append :  
   procedure Append(Object_:IDispatch);safecall;
    // Delete :  
   procedure Delete(Index:OleVariant);safecall;
  end;


// _DynaCollection : 

 _DynaCollectionDisp = dispinterface
   ['{00000513-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Append :  
   procedure Append(Object_:IDispatch);dispid 1610809344;
    // Delete :  
   procedure Delete(Index:OleVariant);dispid 1610809345;
    // Count :  
   property Count:Integer  readonly dispid 1;
  end;


// _ADO : 

 _ADO = interface(IDispatch)
   ['{00000534-0000-0010-8000-00AA006D2EA4}']
   function Get_Properties : Properties; safecall;
    // Properties :  
   property Properties:Properties read Get_Properties;
  end;


// _ADO : 

 _ADODisp = dispinterface
   ['{00000534-0000-0010-8000-00AA006D2EA4}']
    // Properties :  
   property Properties:Properties  readonly dispid 500;
  end;


// Properties : 

 Properties = interface(_Collection)
   ['{00000504-0000-0010-8000-00AA006D2EA4}']
   function Get_Item(Index:OleVariant) : Property_; safecall;
    // Item :  
   property Item[Index:OleVariant]:Property_ read Get_Item; default;
  end;


// Properties : 

 PropertiesDisp = dispinterface
   ['{00000504-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:Property_  readonly dispid 0; default;
  end;


// Property_ : 

 Property_ = interface(IDispatch)
   ['{00000503-0000-0010-8000-00AA006D2EA4}']
   function Get_Value : OleVariant; safecall;
   procedure Set_Value(const pval:OleVariant); safecall;
   function Get_Name : WideString; safecall;
   function Get_Type_ : DataTypeEnum; safecall;
   function Get_Attributes : Integer; safecall;
   procedure Set_Attributes(const plAttributes:Integer); safecall;
    // Value :  
   property Value:OleVariant read Get_Value write Set_Value;
    // Name :  
   property Name:WideString read Get_Name;
    // Type :  
   property Type_:DataTypeEnum read Get_Type_;
    // Attributes :  
   property Attributes:Integer read Get_Attributes write Set_Attributes;
  end;


// Property_ : 

 Property_Disp = dispinterface
   ['{00000503-0000-0010-8000-00AA006D2EA4}']
    // Value :  
   property Value:OleVariant dispid 0;
    // Name :  
   property Name:WideString  readonly dispid 1610743810;
    // Type :  
   property Type_:DataTypeEnum  readonly dispid 1610743811;
    // Attributes :  
   property Attributes:Integer dispid 1610743812;
  end;


// Error : 

 Error = interface(IDispatch)
   ['{00000500-0000-0010-8000-00AA006D2EA4}']
   function Get_Number : Integer; safecall;
   function Get_Source : WideString; safecall;
   function Get_Description : WideString; safecall;
   function Get_HelpFile : WideString; safecall;
   function Get_HelpContext : Integer; safecall;
   function Get_SQLState : WideString; safecall;
   function Get_NativeError : Integer; safecall;
    // Number :  
   property Number:Integer read Get_Number;
    // Source :  
   property Source:WideString read Get_Source;
    // Description :  
   property Description:WideString read Get_Description;
    // HelpFile :  
   property HelpFile:WideString read Get_HelpFile;
    // HelpContext :  
   property HelpContext:Integer read Get_HelpContext;
    // SQLState :  
   property SQLState:WideString read Get_SQLState;
    // NativeError :  
   property NativeError:Integer read Get_NativeError;
  end;


// Error : 

 ErrorDisp = dispinterface
   ['{00000500-0000-0010-8000-00AA006D2EA4}']
    // Number :  
   property Number:Integer  readonly dispid 1;
    // Source :  
   property Source:WideString  readonly dispid 2;
    // Description :  
   property Description:WideString  readonly dispid 0;
    // HelpFile :  
   property HelpFile:WideString  readonly dispid 3;
    // HelpContext :  
   property HelpContext:Integer  readonly dispid 4;
    // SQLState :  
   property SQLState:WideString  readonly dispid 5;
    // NativeError :  
   property NativeError:Integer  readonly dispid 6;
  end;


// Errors : 

 Errors = interface(_Collection)
   ['{00000501-0000-0010-8000-00AA006D2EA4}']
   function Get_Item(Index:OleVariant) : Error; safecall;
    // Clear :  
   procedure Clear;safecall;
    // Item :  
   property Item[Index:OleVariant]:Error read Get_Item; default;
  end;


// Errors : 

 ErrorsDisp = dispinterface
   ['{00000501-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Clear :  
   procedure Clear;dispid 1610809345;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:Error  readonly dispid 0; default;
  end;


// Command15 : 

 Command15 = interface(_ADO)
   ['{00000508-0000-0010-8000-00AA006D2EA4}']
   function Get_ActiveConnection : _Connection; safecall;
   procedure Set_ActiveConnection(const ppvObject:_Connection); safecall;
   procedure Set_ActiveConnection_(const ppvObject:OleVariant); safecall;
   function Get_CommandText : WideString; safecall;
   procedure Set_CommandText(const pbstr:WideString); safecall;
   function Get_CommandTimeout : Integer; safecall;
   procedure Set_CommandTimeout(const pl:Integer); safecall;
   function Get_Prepared : WordBool; safecall;
   procedure Set_Prepared(const pfPrepared:WordBool); safecall;
    // Execute :  
   function Execute(out RecordsAffected:OleVariant;var Parameters:OleVariant;Options:Integer):_Recordset;safecall;
    // CreateParameter :  
   function CreateParameter(Name:WideString;Type_:DataTypeEnum;Direction:ParameterDirectionEnum;Size:ADO_LONGPTR;Value:OleVariant):_Parameter;safecall;
   function Get_Parameters : Parameters; safecall;
   procedure Set_CommandType(const plCmdType:CommandTypeEnum); safecall;
   function Get_CommandType : CommandTypeEnum; safecall;
   function Get_Name : WideString; safecall;
   procedure Set_Name(const pbstrName:WideString); safecall;
    // ActiveConnection :  
   property ActiveConnection:_Connection read Get_ActiveConnection write Set_ActiveConnection;
    // CommandText :  
   property CommandText:WideString read Get_CommandText write Set_CommandText;
    // CommandTimeout :  
   property CommandTimeout:Integer read Get_CommandTimeout write Set_CommandTimeout;
    // Prepared :  
   property Prepared:WordBool read Get_Prepared write Set_Prepared;
    // Parameters :  
   property Parameters:Parameters read Get_Parameters;
    // CommandType :  
   property CommandType:CommandTypeEnum read Get_CommandType write Set_CommandType;
    // Name :  
   property Name:WideString read Get_Name write Set_Name;
  end;


// Command15 : 

 Command15Disp = dispinterface
   ['{00000508-0000-0010-8000-00AA006D2EA4}']
    // Execute :  
   function Execute(out RecordsAffected:OleVariant;var Parameters:OleVariant;Options:Integer):_Recordset;dispid 5;
    // CreateParameter :  
   function CreateParameter(Name:WideString;Type_:DataTypeEnum;Direction:ParameterDirectionEnum;Size:ADO_LONGPTR;Value:OleVariant):_Parameter;dispid 6;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActiveConnection :  
   property ActiveConnection:_Connection dispid 1;
    // CommandText :  
   property CommandText:WideString dispid 2;
    // CommandTimeout :  
   property CommandTimeout:Integer dispid 3;
    // Prepared :  
   property Prepared:WordBool dispid 4;
    // Parameters :  
   property Parameters:Parameters  readonly dispid 0;
    // CommandType :  
   property CommandType:CommandTypeEnum dispid 7;
    // Name :  
   property Name:WideString dispid 8;
  end;


// Connection15 : 

 Connection15 = interface(_ADO)
   ['{00000515-0000-0010-8000-00AA006D2EA4}']
   function Get_ConnectionString : WideString; safecall;
   procedure Set_ConnectionString(const pbstr:WideString); safecall;
   function Get_CommandTimeout : Integer; safecall;
   procedure Set_CommandTimeout(const plTimeout:Integer); safecall;
   function Get_ConnectionTimeout : Integer; safecall;
   procedure Set_ConnectionTimeout(const plTimeout:Integer); safecall;
   function Get_Version : WideString; safecall;
    // Close :  
   procedure Close;safecall;
    // Execute :  
   function Execute(CommandText:WideString;out RecordsAffected:OleVariant;Options:Integer):_Recordset;safecall;
    // BeginTrans :  
   function BeginTrans:Integer;safecall;
    // CommitTrans :  
   procedure CommitTrans;safecall;
    // RollbackTrans :  
   procedure RollbackTrans;safecall;
    // Open :  
   procedure Open(ConnectionString:WideString;UserID:WideString;Password:WideString;Options:Integer);safecall;
   function Get_Errors : Errors; safecall;
   function Get_DefaultDatabase : WideString; safecall;
   procedure Set_DefaultDatabase(const pbstr:WideString); safecall;
   function Get_IsolationLevel : IsolationLevelEnum; safecall;
   procedure Set_IsolationLevel(const Level:IsolationLevelEnum); safecall;
   function Get_Attributes : Integer; safecall;
   procedure Set_Attributes(const plAttr:Integer); safecall;
   function Get_CursorLocation : CursorLocationEnum; safecall;
   procedure Set_CursorLocation(const plCursorLoc:CursorLocationEnum); safecall;
   function Get_Mode : ConnectModeEnum; safecall;
   procedure Set_Mode(const plMode:ConnectModeEnum); safecall;
   function Get_Provider : WideString; safecall;
   procedure Set_Provider(const pbstr:WideString); safecall;
   function Get_State : Integer; safecall;
    // OpenSchema :  
   function OpenSchema(Schema:SchemaEnum;Restrictions:OleVariant;SchemaID:OleVariant):_Recordset;safecall;
    // ConnectionString :  
   property ConnectionString:WideString read Get_ConnectionString write Set_ConnectionString;
    // CommandTimeout :  
   property CommandTimeout:Integer read Get_CommandTimeout write Set_CommandTimeout;
    // ConnectionTimeout :  
   property ConnectionTimeout:Integer read Get_ConnectionTimeout write Set_ConnectionTimeout;
    // Version :  
   property Version:WideString read Get_Version;
    // Errors :  
   property Errors:Errors read Get_Errors;
    // DefaultDatabase :  
   property DefaultDatabase:WideString read Get_DefaultDatabase write Set_DefaultDatabase;
    // IsolationLevel :  
   property IsolationLevel:IsolationLevelEnum read Get_IsolationLevel write Set_IsolationLevel;
    // Attributes :  
   property Attributes:Integer read Get_Attributes write Set_Attributes;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum read Get_CursorLocation write Set_CursorLocation;
    // Mode :  
   property Mode:ConnectModeEnum read Get_Mode write Set_Mode;
    // Provider :  
   property Provider:WideString read Get_Provider write Set_Provider;
    // State :  
   property State:Integer read Get_State;
  end;


// Connection15 : 

 Connection15Disp = dispinterface
   ['{00000515-0000-0010-8000-00AA006D2EA4}']
    // Close :  
   procedure Close;dispid 5;
    // Execute :  
   function Execute(CommandText:WideString;out RecordsAffected:OleVariant;Options:Integer):_Recordset;dispid 6;
    // BeginTrans :  
   function BeginTrans:Integer;dispid 7;
    // CommitTrans :  
   procedure CommitTrans;dispid 8;
    // RollbackTrans :  
   procedure RollbackTrans;dispid 9;
    // Open :  
   procedure Open(ConnectionString:WideString;UserID:WideString;Password:WideString;Options:Integer);dispid 10;
    // OpenSchema :  
   function OpenSchema(Schema:SchemaEnum;Restrictions:OleVariant;SchemaID:OleVariant):_Recordset;dispid 19;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ConnectionString :  
   property ConnectionString:WideString dispid 0;
    // CommandTimeout :  
   property CommandTimeout:Integer dispid 2;
    // ConnectionTimeout :  
   property ConnectionTimeout:Integer dispid 3;
    // Version :  
   property Version:WideString  readonly dispid 4;
    // Errors :  
   property Errors:Errors  readonly dispid 11;
    // DefaultDatabase :  
   property DefaultDatabase:WideString dispid 12;
    // IsolationLevel :  
   property IsolationLevel:IsolationLevelEnum dispid 13;
    // Attributes :  
   property Attributes:Integer dispid 14;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 15;
    // Mode :  
   property Mode:ConnectModeEnum dispid 16;
    // Provider :  
   property Provider:WideString dispid 17;
    // State :  
   property State:Integer  readonly dispid 18;
  end;


// _Connection : 

 _Connection = interface(Connection15)
   ['{00000550-0000-0010-8000-00AA006D2EA4}']
    // Cancel :  
   procedure Cancel;safecall;
  end;

// _Connection : 

 _ConnectionDisp = dispinterface
   ['{00000550-0000-0010-8000-00AA006D2EA4}']
    // Close :  
   procedure Close;dispid 5;
    // Execute :  
   function Execute(CommandText:WideString;out RecordsAffected:OleVariant;Options:Integer):_Recordset;dispid 6;
    // BeginTrans :  
   function BeginTrans:Integer;dispid 7;
    // CommitTrans :  
   procedure CommitTrans;dispid 8;
    // RollbackTrans :  
   procedure RollbackTrans;dispid 9;
    // Open :  
   procedure Open(ConnectionString:WideString;UserID:WideString;Password:WideString;Options:Integer);dispid 10;
    // OpenSchema :  
   function OpenSchema(Schema:SchemaEnum;Restrictions:OleVariant;SchemaID:OleVariant):_Recordset;dispid 19;
    // Cancel :  
   procedure Cancel;dispid 21;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ConnectionString :  
   property ConnectionString:WideString dispid 0;
    // CommandTimeout :  
   property CommandTimeout:Integer dispid 2;
    // ConnectionTimeout :  
   property ConnectionTimeout:Integer dispid 3;
    // Version :  
   property Version:WideString  readonly dispid 4;
    // Errors :  
   property Errors:Errors  readonly dispid 11;
    // DefaultDatabase :  
   property DefaultDatabase:WideString dispid 12;
    // IsolationLevel :  
   property IsolationLevel:IsolationLevelEnum dispid 13;
    // Attributes :  
   property Attributes:Integer dispid 14;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 15;
    // Mode :  
   property Mode:ConnectModeEnum dispid 16;
    // Provider :  
   property Provider:WideString dispid 17;
    // State :  
   property State:Integer  readonly dispid 18;
  end;


// Recordset15 : 

 Recordset15 = interface(_ADO)
   ['{0000050E-0000-0010-8000-00AA006D2EA4}']
   function Get_AbsolutePosition : PositionEnum_Param; safecall;
   procedure Set_AbsolutePosition(const pl:PositionEnum_Param); safecall;
   procedure Set_ActiveConnection(const pvar:IDispatch); safecall;
   procedure Set_ActiveConnection_(const pvar:OleVariant); safecall;
   function Get_ActiveConnection : OleVariant; safecall;
   function Get_BOF : WordBool; safecall;
   function Get_Bookmark : OleVariant; safecall;
   procedure Set_Bookmark(const pvBookmark:OleVariant); safecall;
   function Get_CacheSize : Integer; safecall;
   procedure Set_CacheSize(const pl:Integer); safecall;
   function Get_CursorType : CursorTypeEnum; safecall;
   procedure Set_CursorType(const plCursorType:CursorTypeEnum); safecall;
   function Get_EOF_ : WordBool; safecall;
   function Get_Fields : Fields; safecall;
   function Get_LockType : LockTypeEnum; safecall;
   procedure Set_LockType(const plLockType:LockTypeEnum); safecall;
   function Get_MaxRecords : ADO_LONGPTR; safecall;
   procedure Set_MaxRecords(const plMaxRecords:ADO_LONGPTR); safecall;
   function Get_RecordCount : ADO_LONGPTR; safecall;
   procedure Set_Source(const pvSource:IDispatch); safecall;
   procedure Set_Source_(const pvSource:WideString); safecall;
   function Get_Source : OleVariant; safecall;
    // AddNew :  
   procedure AddNew(FieldList:OleVariant;Values:OleVariant);safecall;
    // CancelUpdate :  
   procedure CancelUpdate;safecall;
    // Close :  
   procedure Close;safecall;
    // Delete :  
   procedure Delete(AffectRecords:AffectEnum);safecall;
    // GetRows :  
   function GetRows(Rows:Integer;Start:OleVariant;Fields:OleVariant):OleVariant;safecall;
    // Move :  
   procedure Move(NumRecords:ADO_LONGPTR;Start:OleVariant);safecall;
    // MoveNext :  
   procedure MoveNext;safecall;
    // MovePrevious :  
   procedure MovePrevious;safecall;
    // MoveFirst :  
   procedure MoveFirst;safecall;
    // MoveLast :  
   procedure MoveLast;safecall;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;CursorType:CursorTypeEnum;LockType:LockTypeEnum;Options:Integer);safecall;
    // Requery :  
   procedure Requery(Options:Integer);safecall;
    // _xResync :  
   procedure _xResync(AffectRecords:AffectEnum);safecall;
    // Update :  
   procedure Update(Fields:OleVariant;Values:OleVariant);safecall;
   function Get_AbsolutePage : PositionEnum_Param; safecall;
   procedure Set_AbsolutePage(const pl:PositionEnum_Param); safecall;
   function Get_EditMode : EditModeEnum; safecall;
   function Get_Filter : OleVariant; safecall;
   procedure Set_Filter(const Criteria:OleVariant); safecall;
   function Get_PageCount : ADO_LONGPTR; safecall;
   function Get_PageSize : Integer; safecall;
   procedure Set_PageSize(const pl:Integer); safecall;
   function Get_Sort : WideString; safecall;
   procedure Set_Sort(const Criteria:WideString); safecall;
   function Get_Status : Integer; safecall;
   function Get_State : Integer; safecall;
    // _xClone :  
   function _xClone:_Recordset;safecall;
    // UpdateBatch :  
   procedure UpdateBatch(AffectRecords:AffectEnum);safecall;
    // CancelBatch :  
   procedure CancelBatch(AffectRecords:AffectEnum);safecall;
   function Get_CursorLocation : CursorLocationEnum; safecall;
   procedure Set_CursorLocation(const plCursorLoc:CursorLocationEnum); safecall;
    // NextRecordset :  
   function NextRecordset(out RecordsAffected:OleVariant):_Recordset;safecall;
    // Supports :  
   function Supports(CursorOptions:CursorOptionEnum):WordBool;safecall;
   function Get_Collect(Index:OleVariant) : OleVariant; safecall;
   procedure Set_Collect(const Index:OleVariant; const parCollect:OleVariant); safecall;
   function Get_MarshalOptions : MarshalOptionsEnum; safecall;
   procedure Set_MarshalOptions(const peMarshal:MarshalOptionsEnum); safecall;
    // Find :  
   procedure Find(Criteria:WideString;SkipRecords:ADO_LONGPTR;SearchDirection:SearchDirectionEnum;Start:OleVariant);safecall;
    // AbsolutePosition :  
   property AbsolutePosition:PositionEnum_Param read Get_AbsolutePosition write Set_AbsolutePosition;
    // ActiveConnection :  
   property ActiveConnection:OleVariant read Get_ActiveConnection write Set_ActiveConnection_;
    // BOF :  
   property BOF:WordBool read Get_BOF;
    // Bookmark :  
   property Bookmark:OleVariant read Get_Bookmark write Set_Bookmark;
    // CacheSize :  
   property CacheSize:Integer read Get_CacheSize write Set_CacheSize;
    // CursorType :  
   property CursorType:CursorTypeEnum read Get_CursorType write Set_CursorType;
    // EOF :  
   property EOF_:WordBool read Get_EOF_;
    // Fields :  
   property Fields:Fields read Get_Fields;
    // LockType :  
   property LockType:LockTypeEnum read Get_LockType write Set_LockType;
    // MaxRecords :  
   property MaxRecords:ADO_LONGPTR read Get_MaxRecords write Set_MaxRecords;
    // RecordCount :  
   property RecordCount:ADO_LONGPTR read Get_RecordCount;
    // AbsolutePage :  
   property AbsolutePage:PositionEnum_Param read Get_AbsolutePage write Set_AbsolutePage;
    // EditMode :  
   property EditMode:EditModeEnum read Get_EditMode;
    // Filter :  
   property Filter:OleVariant read Get_Filter write Set_Filter;
    // PageCount :  
   property PageCount:ADO_LONGPTR read Get_PageCount;
    // PageSize :  
   property PageSize:Integer read Get_PageSize write Set_PageSize;
    // Sort :  
   property Sort:WideString read Get_Sort write Set_Sort;
    // Status :  
   property Status:Integer read Get_Status;
    // State :  
   property State:Integer read Get_State;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum read Get_CursorLocation write Set_CursorLocation;
    // Collect :  
   property Collect[Index:OleVariant]:OleVariant read Get_Collect write Set_Collect;
    // MarshalOptions :  
   property MarshalOptions:MarshalOptionsEnum read Get_MarshalOptions write Set_MarshalOptions;
  end;


// Recordset15 : 

 Recordset15Disp = dispinterface
   ['{0000050E-0000-0010-8000-00AA006D2EA4}']
    // AddNew :  
   procedure AddNew(FieldList:OleVariant;Values:OleVariant);dispid 1012;
    // CancelUpdate :  
   procedure CancelUpdate;dispid 1013;
    // Close :  
   procedure Close;dispid 1014;
    // Delete :  
   procedure Delete(AffectRecords:AffectEnum);dispid 1015;
    // GetRows :  
   function GetRows(Rows:Integer;Start:OleVariant;Fields:OleVariant):OleVariant;dispid 1016;
    // Move :  
   procedure Move(NumRecords:ADO_LONGPTR;Start:OleVariant);dispid 1017;
    // MoveNext :  
   procedure MoveNext;dispid 1018;
    // MovePrevious :  
   procedure MovePrevious;dispid 1019;
    // MoveFirst :  
   procedure MoveFirst;dispid 1020;
    // MoveLast :  
   procedure MoveLast;dispid 1021;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;CursorType:CursorTypeEnum;LockType:LockTypeEnum;Options:Integer);dispid 1022;
    // Requery :  
   procedure Requery(Options:Integer);dispid 1023;
    // _xResync :  
   procedure _xResync(AffectRecords:AffectEnum);dispid 1610809378;
    // Update :  
   procedure Update(Fields:OleVariant;Values:OleVariant);dispid 1025;
    // _xClone :  
   function _xClone:_Recordset;dispid 1610809392;
    // UpdateBatch :  
   procedure UpdateBatch(AffectRecords:AffectEnum);dispid 1035;
    // CancelBatch :  
   procedure CancelBatch(AffectRecords:AffectEnum);dispid 1049;
    // NextRecordset :  
   function NextRecordset(out RecordsAffected:OleVariant):_Recordset;dispid 1052;
    // Supports :  
   function Supports(CursorOptions:CursorOptionEnum):WordBool;dispid 1036;
    // Find :  
   procedure Find(Criteria:WideString;SkipRecords:ADO_LONGPTR;SearchDirection:SearchDirectionEnum;Start:OleVariant);dispid 1058;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // AbsolutePosition :  
   property AbsolutePosition:PositionEnum_Param dispid 1000;
    // ActiveConnection :  
   property ActiveConnection:IDispatch dispid 1001;
    // BOF :  
   property BOF:WordBool  readonly dispid 1002;
    // Bookmark :  
   property Bookmark:OleVariant dispid 1003;
    // CacheSize :  
   property CacheSize:Integer dispid 1004;
    // CursorType :  
   property CursorType:CursorTypeEnum dispid 1005;
    // EOF :  
   property EOF_:WordBool  readonly dispid 1006;
    // Fields :  
   property Fields:Fields  readonly dispid 0;
    // LockType :  
   property LockType:LockTypeEnum dispid 1008;
    // MaxRecords :  
   property MaxRecords:ADO_LONGPTR dispid 1009;
    // RecordCount :  
   property RecordCount:ADO_LONGPTR  readonly dispid 1010;
    // Source :  
   property Source:IDispatch dispid 1011;
    // AbsolutePage :  
   property AbsolutePage:PositionEnum_Param dispid 1047;
    // EditMode :  
   property EditMode:EditModeEnum  readonly dispid 1026;
    // Filter :  
   property Filter:OleVariant dispid 1030;
    // PageCount :  
   property PageCount:ADO_LONGPTR  readonly dispid 1050;
    // PageSize :  
   property PageSize:Integer dispid 1048;
    // Sort :  
   property Sort:WideString dispid 1031;
    // Status :  
   property Status:Integer  readonly dispid 1029;
    // State :  
   property State:Integer  readonly dispid 1054;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 1051;
    // Collect :  
   property Collect[Index:OleVariant]:OleVariant dispid -8;
    // MarshalOptions :  
   property MarshalOptions:MarshalOptionsEnum dispid 1053;
  end;


// Recordset20 : 

 Recordset20 = interface(Recordset15)
   ['{0000054F-0000-0010-8000-00AA006D2EA4}']
    // Cancel :  
   procedure Cancel;safecall;
   function Get_DataSource : IUnknown; safecall;
   procedure Set_DataSource(const ppunkDataSource:IUnknown); safecall;
    // _xSave :  
   procedure _xSave(FileName:WideString;PersistFormat:PersistFormatEnum);safecall;
   function Get_ActiveCommand : IDispatch; safecall;
   procedure Set_StayInSync(const pbStayInSync:WordBool); safecall;
   function Get_StayInSync : WordBool; safecall;
    // GetString :  
   function GetString(StringFormat:StringFormatEnum;NumRows:Integer;ColumnDelimeter:WideString;RowDelimeter:WideString;NullExpr:WideString):WideString;safecall;
   function Get_DataMember : WideString; safecall;
   procedure Set_DataMember(const pbstrDataMember:WideString); safecall;
    // CompareBookmarks :  
   function CompareBookmarks(Bookmark1:OleVariant;Bookmark2:OleVariant):CompareEnum;safecall;
    // Clone :  
   function Clone(LockType:LockTypeEnum):_Recordset;safecall;
    // Resync :  
   procedure Resync(AffectRecords:AffectEnum;ResyncValues:ResyncEnum);safecall;
    // DataSource :  
   property DataSource:IUnknown read Get_DataSource write Set_DataSource;
    // ActiveCommand :  
   property ActiveCommand:IDispatch read Get_ActiveCommand;
    // StayInSync :  
   property StayInSync:WordBool read Get_StayInSync write Set_StayInSync;
    // DataMember :  
   property DataMember:WideString read Get_DataMember write Set_DataMember;
  end;

// Recordset20 : 

 Recordset20Disp = dispinterface
   ['{0000054F-0000-0010-8000-00AA006D2EA4}']
    // AddNew :  
   procedure AddNew(FieldList:OleVariant;Values:OleVariant);dispid 1012;
    // CancelUpdate :  
   procedure CancelUpdate;dispid 1013;
    // Close :  
   procedure Close;dispid 1014;
    // Delete :  
   procedure Delete(AffectRecords:AffectEnum);dispid 1015;
    // GetRows :  
   function GetRows(Rows:Integer;Start:OleVariant;Fields:OleVariant):OleVariant;dispid 1016;
    // Move :  
   procedure Move(NumRecords:ADO_LONGPTR;Start:OleVariant);dispid 1017;
    // MoveNext :  
   procedure MoveNext;dispid 1018;
    // MovePrevious :  
   procedure MovePrevious;dispid 1019;
    // MoveFirst :  
   procedure MoveFirst;dispid 1020;
    // MoveLast :  
   procedure MoveLast;dispid 1021;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;CursorType:CursorTypeEnum;LockType:LockTypeEnum;Options:Integer);dispid 1022;
    // Requery :  
   procedure Requery(Options:Integer);dispid 1023;
    // _xResync :  
   procedure _xResync(AffectRecords:AffectEnum);dispid 1610809378;
    // Update :  
   procedure Update(Fields:OleVariant;Values:OleVariant);dispid 1025;
    // _xClone :  
   function _xClone:_Recordset;dispid 1610809392;
    // UpdateBatch :  
   procedure UpdateBatch(AffectRecords:AffectEnum);dispid 1035;
    // CancelBatch :  
   procedure CancelBatch(AffectRecords:AffectEnum);dispid 1049;
    // NextRecordset :  
   function NextRecordset(out RecordsAffected:OleVariant):_Recordset;dispid 1052;
    // Supports :  
   function Supports(CursorOptions:CursorOptionEnum):WordBool;dispid 1036;
    // Find :  
   procedure Find(Criteria:WideString;SkipRecords:ADO_LONGPTR;SearchDirection:SearchDirectionEnum;Start:OleVariant);dispid 1058;
    // Cancel :  
   procedure Cancel;dispid 1055;
    // _xSave :  
   procedure _xSave(FileName:WideString;PersistFormat:PersistFormatEnum);dispid 1610874883;
    // GetString :  
   function GetString(StringFormat:StringFormatEnum;NumRows:Integer;ColumnDelimeter:WideString;RowDelimeter:WideString;NullExpr:WideString):WideString;dispid 1062;
    // CompareBookmarks :  
   function CompareBookmarks(Bookmark1:OleVariant;Bookmark2:OleVariant):CompareEnum;dispid 1065;
    // Clone :  
   function Clone(LockType:LockTypeEnum):_Recordset;dispid 1034;
    // Resync :  
   procedure Resync(AffectRecords:AffectEnum;ResyncValues:ResyncEnum);dispid 1024;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // AbsolutePosition :  
   property AbsolutePosition:PositionEnum_Param dispid 1000;
    // ActiveConnection :  
   property ActiveConnection:IDispatch dispid 1001;
    // BOF :  
   property BOF:WordBool  readonly dispid 1002;
    // Bookmark :  
   property Bookmark:OleVariant dispid 1003;
    // CacheSize :  
   property CacheSize:Integer dispid 1004;
    // CursorType :  
   property CursorType:CursorTypeEnum dispid 1005;
    // EOF :  
   property EOF_:WordBool  readonly dispid 1006;
    // Fields :  
   property Fields:Fields  readonly dispid 0;
    // LockType :  
   property LockType:LockTypeEnum dispid 1008;
    // MaxRecords :  
   property MaxRecords:ADO_LONGPTR dispid 1009;
    // RecordCount :  
   property RecordCount:ADO_LONGPTR  readonly dispid 1010;
    // Source :  
   property Source:IDispatch dispid 1011;
    // AbsolutePage :  
   property AbsolutePage:PositionEnum_Param dispid 1047;
    // EditMode :  
   property EditMode:EditModeEnum  readonly dispid 1026;
    // Filter :  
   property Filter:OleVariant dispid 1030;
    // PageCount :  
   property PageCount:ADO_LONGPTR  readonly dispid 1050;
    // PageSize :  
   property PageSize:Integer dispid 1048;
    // Sort :  
   property Sort:WideString dispid 1031;
    // Status :  
   property Status:Integer  readonly dispid 1029;
    // State :  
   property State:Integer  readonly dispid 1054;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 1051;
    // Collect :  
   property Collect[Index:OleVariant]:OleVariant dispid -8;
    // MarshalOptions :  
   property MarshalOptions:MarshalOptionsEnum dispid 1053;
    // DataSource :  
   property DataSource:IUnknown dispid 1056;
    // ActiveCommand :  
   property ActiveCommand:IDispatch  readonly dispid 1061;
    // StayInSync :  
   property StayInSync:WordBool dispid 1063;
    // DataMember :  
   property DataMember:WideString dispid 1064;
  end;


// Recordset21 : 

 Recordset21 = interface(Recordset20)
   ['{00000555-0000-0010-8000-00AA006D2EA4}']
    // Seek :  
   procedure Seek(KeyValues:OleVariant;SeekOption:SeekEnum);safecall;
   procedure Set_Index(const pbstrIndex:WideString); safecall;
   function Get_Index : WideString; safecall;
    // Index :  
   property Index:WideString read Get_Index write Set_Index;
  end;

// Recordset21 : 

 Recordset21Disp = dispinterface
   ['{00000555-0000-0010-8000-00AA006D2EA4}']
    // AddNew :  
   procedure AddNew(FieldList:OleVariant;Values:OleVariant);dispid 1012;
    // CancelUpdate :  
   procedure CancelUpdate;dispid 1013;
    // Close :  
   procedure Close;dispid 1014;
    // Delete :  
   procedure Delete(AffectRecords:AffectEnum);dispid 1015;
    // GetRows :  
   function GetRows(Rows:Integer;Start:OleVariant;Fields:OleVariant):OleVariant;dispid 1016;
    // Move :  
   procedure Move(NumRecords:ADO_LONGPTR;Start:OleVariant);dispid 1017;
    // MoveNext :  
   procedure MoveNext;dispid 1018;
    // MovePrevious :  
   procedure MovePrevious;dispid 1019;
    // MoveFirst :  
   procedure MoveFirst;dispid 1020;
    // MoveLast :  
   procedure MoveLast;dispid 1021;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;CursorType:CursorTypeEnum;LockType:LockTypeEnum;Options:Integer);dispid 1022;
    // Requery :  
   procedure Requery(Options:Integer);dispid 1023;
    // _xResync :  
   procedure _xResync(AffectRecords:AffectEnum);dispid 1610809378;
    // Update :  
   procedure Update(Fields:OleVariant;Values:OleVariant);dispid 1025;
    // _xClone :  
   function _xClone:_Recordset;dispid 1610809392;
    // UpdateBatch :  
   procedure UpdateBatch(AffectRecords:AffectEnum);dispid 1035;
    // CancelBatch :  
   procedure CancelBatch(AffectRecords:AffectEnum);dispid 1049;
    // NextRecordset :  
   function NextRecordset(out RecordsAffected:OleVariant):_Recordset;dispid 1052;
    // Supports :  
   function Supports(CursorOptions:CursorOptionEnum):WordBool;dispid 1036;
    // Find :  
   procedure Find(Criteria:WideString;SkipRecords:ADO_LONGPTR;SearchDirection:SearchDirectionEnum;Start:OleVariant);dispid 1058;
    // Cancel :  
   procedure Cancel;dispid 1055;
    // _xSave :  
   procedure _xSave(FileName:WideString;PersistFormat:PersistFormatEnum);dispid 1610874883;
    // GetString :  
   function GetString(StringFormat:StringFormatEnum;NumRows:Integer;ColumnDelimeter:WideString;RowDelimeter:WideString;NullExpr:WideString):WideString;dispid 1062;
    // CompareBookmarks :  
   function CompareBookmarks(Bookmark1:OleVariant;Bookmark2:OleVariant):CompareEnum;dispid 1065;
    // Clone :  
   function Clone(LockType:LockTypeEnum):_Recordset;dispid 1034;
    // Resync :  
   procedure Resync(AffectRecords:AffectEnum;ResyncValues:ResyncEnum);dispid 1024;
    // Seek :  
   procedure Seek(KeyValues:OleVariant;SeekOption:SeekEnum);dispid 1066;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // AbsolutePosition :  
   property AbsolutePosition:PositionEnum_Param dispid 1000;
    // ActiveConnection :  
   property ActiveConnection:IDispatch dispid 1001;
    // BOF :  
   property BOF:WordBool  readonly dispid 1002;
    // Bookmark :  
   property Bookmark:OleVariant dispid 1003;
    // CacheSize :  
   property CacheSize:Integer dispid 1004;
    // CursorType :  
   property CursorType:CursorTypeEnum dispid 1005;
    // EOF :  
   property EOF_:WordBool  readonly dispid 1006;
    // Fields :  
   property Fields:Fields  readonly dispid 0;
    // LockType :  
   property LockType:LockTypeEnum dispid 1008;
    // MaxRecords :  
   property MaxRecords:ADO_LONGPTR dispid 1009;
    // RecordCount :  
   property RecordCount:ADO_LONGPTR  readonly dispid 1010;
    // Source :  
   property Source:IDispatch dispid 1011;
    // AbsolutePage :  
   property AbsolutePage:PositionEnum_Param dispid 1047;
    // EditMode :  
   property EditMode:EditModeEnum  readonly dispid 1026;
    // Filter :  
   property Filter:OleVariant dispid 1030;
    // PageCount :  
   property PageCount:ADO_LONGPTR  readonly dispid 1050;
    // PageSize :  
   property PageSize:Integer dispid 1048;
    // Sort :  
   property Sort:WideString dispid 1031;
    // Status :  
   property Status:Integer  readonly dispid 1029;
    // State :  
   property State:Integer  readonly dispid 1054;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 1051;
    // Collect :  
   property Collect[Index:OleVariant]:OleVariant dispid -8;
    // MarshalOptions :  
   property MarshalOptions:MarshalOptionsEnum dispid 1053;
    // DataSource :  
   property DataSource:IUnknown dispid 1056;
    // ActiveCommand :  
   property ActiveCommand:IDispatch  readonly dispid 1061;
    // StayInSync :  
   property StayInSync:WordBool dispid 1063;
    // DataMember :  
   property DataMember:WideString dispid 1064;
    // Index :  
   property Index:WideString dispid 1067;
  end;


// _Recordset : 

 _Recordset = interface(Recordset21)
   ['{00000556-0000-0010-8000-00AA006D2EA4}']
    // Save :  
   procedure Save(Destination:OleVariant;PersistFormat:PersistFormatEnum);safecall;
  end;

// _Recordset : 

 _RecordsetDisp = dispinterface
   ['{00000556-0000-0010-8000-00AA006D2EA4}']
    // AddNew :  
   procedure AddNew(FieldList:OleVariant;Values:OleVariant);dispid 1012;
    // CancelUpdate :  
   procedure CancelUpdate;dispid 1013;
    // Close :  
   procedure Close;dispid 1014;
    // Delete :  
   procedure Delete(AffectRecords:AffectEnum);dispid 1015;
    // GetRows :  
   function GetRows(Rows:Integer;Start:OleVariant;Fields:OleVariant):OleVariant;dispid 1016;
    // Move :  
   procedure Move(NumRecords:ADO_LONGPTR;Start:OleVariant);dispid 1017;
    // MoveNext :  
   procedure MoveNext;dispid 1018;
    // MovePrevious :  
   procedure MovePrevious;dispid 1019;
    // MoveFirst :  
   procedure MoveFirst;dispid 1020;
    // MoveLast :  
   procedure MoveLast;dispid 1021;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;CursorType:CursorTypeEnum;LockType:LockTypeEnum;Options:Integer);dispid 1022;
    // Requery :  
   procedure Requery(Options:Integer);dispid 1023;
    // _xResync :  
   procedure _xResync(AffectRecords:AffectEnum);dispid 1610809378;
    // Update :  
   procedure Update(Fields:OleVariant;Values:OleVariant);dispid 1025;
    // _xClone :  
   function _xClone:_Recordset;dispid 1610809392;
    // UpdateBatch :  
   procedure UpdateBatch(AffectRecords:AffectEnum);dispid 1035;
    // CancelBatch :  
   procedure CancelBatch(AffectRecords:AffectEnum);dispid 1049;
    // NextRecordset :  
   function NextRecordset(out RecordsAffected:OleVariant):_Recordset;dispid 1052;
    // Supports :  
   function Supports(CursorOptions:CursorOptionEnum):WordBool;dispid 1036;
    // Find :  
   procedure Find(Criteria:WideString;SkipRecords:ADO_LONGPTR;SearchDirection:SearchDirectionEnum;Start:OleVariant);dispid 1058;
    // Cancel :  
   procedure Cancel;dispid 1055;
    // _xSave :  
   procedure _xSave(FileName:WideString;PersistFormat:PersistFormatEnum);dispid 1610874883;
    // GetString :  
   function GetString(StringFormat:StringFormatEnum;NumRows:Integer;ColumnDelimeter:WideString;RowDelimeter:WideString;NullExpr:WideString):WideString;dispid 1062;
    // CompareBookmarks :  
   function CompareBookmarks(Bookmark1:OleVariant;Bookmark2:OleVariant):CompareEnum;dispid 1065;
    // Clone :  
   function Clone(LockType:LockTypeEnum):_Recordset;dispid 1034;
    // Resync :  
   procedure Resync(AffectRecords:AffectEnum;ResyncValues:ResyncEnum);dispid 1024;
    // Seek :  
   procedure Seek(KeyValues:OleVariant;SeekOption:SeekEnum);dispid 1066;
    // Save :  
   procedure Save(Destination:OleVariant;PersistFormat:PersistFormatEnum);dispid 1057;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // AbsolutePosition :  
   property AbsolutePosition:PositionEnum_Param dispid 1000;
    // ActiveConnection :  
   property ActiveConnection:IDispatch dispid 1001;
    // BOF :  
   property BOF:WordBool  readonly dispid 1002;
    // Bookmark :  
   property Bookmark:OleVariant dispid 1003;
    // CacheSize :  
   property CacheSize:Integer dispid 1004;
    // CursorType :  
   property CursorType:CursorTypeEnum dispid 1005;
    // EOF :  
   property EOF_:WordBool  readonly dispid 1006;
    // Fields :  
   property Fields:Fields  readonly dispid 0;
    // LockType :  
   property LockType:LockTypeEnum dispid 1008;
    // MaxRecords :  
   property MaxRecords:ADO_LONGPTR dispid 1009;
    // RecordCount :  
   property RecordCount:ADO_LONGPTR  readonly dispid 1010;
    // Source :  
   property Source:IDispatch dispid 1011;
    // AbsolutePage :  
   property AbsolutePage:PositionEnum_Param dispid 1047;
    // EditMode :  
   property EditMode:EditModeEnum  readonly dispid 1026;
    // Filter :  
   property Filter:OleVariant dispid 1030;
    // PageCount :  
   property PageCount:ADO_LONGPTR  readonly dispid 1050;
    // PageSize :  
   property PageSize:Integer dispid 1048;
    // Sort :  
   property Sort:WideString dispid 1031;
    // Status :  
   property Status:Integer  readonly dispid 1029;
    // State :  
   property State:Integer  readonly dispid 1054;
    // CursorLocation :  
   property CursorLocation:CursorLocationEnum dispid 1051;
    // Collect :  
   property Collect[Index:OleVariant]:OleVariant dispid -8;
    // MarshalOptions :  
   property MarshalOptions:MarshalOptionsEnum dispid 1053;
    // DataSource :  
   property DataSource:IUnknown dispid 1056;
    // ActiveCommand :  
   property ActiveCommand:IDispatch  readonly dispid 1061;
    // StayInSync :  
   property StayInSync:WordBool dispid 1063;
    // DataMember :  
   property DataMember:WideString dispid 1064;
    // Index :  
   property Index:WideString dispid 1067;
  end;


// Fields15 : 

 Fields15 = interface(_Collection)
   ['{00000506-0000-0010-8000-00AA006D2EA4}']
   function Get_Item(Index:OleVariant) : Field; safecall;
    // Item :  
   property Item[Index:OleVariant]:Field read Get_Item; default;
  end;


// Fields15 : 

 Fields15Disp = dispinterface
   ['{00000506-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:Field  readonly dispid 0; default;
  end;


// Fields20 : 

 Fields20 = interface(Fields15)
   ['{0000054D-0000-0010-8000-00AA006D2EA4}']
    // _Append :  
   procedure _Append(Name:WideString;Type_:DataTypeEnum;DefinedSize:ADO_LONGPTR;Attrib:FieldAttributeEnum);safecall;
    // Delete :  
   procedure Delete(Index:OleVariant);safecall;
  end;

// Fields20 : 

 Fields20Disp = dispinterface
   ['{0000054D-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // _Append :  
   procedure _Append(Name:WideString;Type_:DataTypeEnum;DefinedSize:ADO_LONGPTR;Attrib:FieldAttributeEnum);dispid 1610874880;
    // Delete :  
   procedure Delete(Index:OleVariant);dispid 4;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:Field  readonly dispid 0; default;
  end;


// Fields : 

 Fields = interface(Fields20)
   ['{00000564-0000-0010-8000-00AA006D2EA4}']
    // Append :  
   procedure Append(Name:WideString;Type_:DataTypeEnum;DefinedSize:ADO_LONGPTR;Attrib:FieldAttributeEnum;FieldValue:OleVariant);safecall;
    // Update :  
   procedure Update;safecall;
    // Resync :  
   procedure Resync(ResyncValues:ResyncEnum);safecall;
    // CancelUpdate :  
   procedure CancelUpdate;safecall;
  end;

// Fields : 

 FieldsDisp = dispinterface
   ['{00000564-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // _Append :  
   procedure _Append(Name:WideString;Type_:DataTypeEnum;DefinedSize:ADO_LONGPTR;Attrib:FieldAttributeEnum);dispid 1610874880;
    // Delete :  
   procedure Delete(Index:OleVariant);dispid 4;
    // Append :  
   procedure Append(Name:WideString;Type_:DataTypeEnum;DefinedSize:ADO_LONGPTR;Attrib:FieldAttributeEnum;FieldValue:OleVariant);dispid 3;
    // Update :  
   procedure Update;dispid 5;
    // Resync :  
   procedure Resync(ResyncValues:ResyncEnum);dispid 6;
    // CancelUpdate :  
   procedure CancelUpdate;dispid 7;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:Field  readonly dispid 0; default;
  end;


// Field20 : 

 Field20 = interface(_ADO)
   ['{0000054C-0000-0010-8000-00AA006D2EA4}']
   function Get_ActualSize : ADO_LONGPTR; safecall;
   function Get_Attributes : Integer; safecall;
   function Get_DefinedSize : ADO_LONGPTR; safecall;
   function Get_Name : WideString; safecall;
   function Get_Type_ : DataTypeEnum; safecall;
   function Get_Value : OleVariant; safecall;
   procedure Set_Value(const pvar:OleVariant); safecall;
   function Get_Precision : Byte; safecall;
   function Get_NumericScale : Byte; safecall;
    // AppendChunk :  
   procedure AppendChunk(Data:OleVariant);safecall;
    // GetChunk :  
   function GetChunk(Length:Integer):OleVariant;safecall;
   function Get_OriginalValue : OleVariant; safecall;
   function Get_UnderlyingValue : OleVariant; safecall;
   function Get_DataFormat : IUnknown; safecall;
   procedure Set_DataFormat(const ppiDF:IUnknown); safecall;
   procedure Set_Precision(const pbPrecision:Byte); safecall;
   procedure Set_NumericScale(const pbNumericScale:Byte); safecall;
   procedure Set_Type_(const pDataType:DataTypeEnum); safecall;
   procedure Set_DefinedSize(const pl:ADO_LONGPTR); safecall;
   procedure Set_Attributes(const pl:Integer); safecall;
    // ActualSize :  
   property ActualSize:ADO_LONGPTR read Get_ActualSize;
    // Attributes :  
   property Attributes:Integer read Get_Attributes write Set_Attributes;
    // DefinedSize :  
   property DefinedSize:ADO_LONGPTR read Get_DefinedSize write Set_DefinedSize;
    // Name :  
   property Name:WideString read Get_Name;
    // Type :  
   property Type_:DataTypeEnum read Get_Type_ write Set_Type_;
    // Value :  
   property Value:OleVariant read Get_Value write Set_Value;
    // Precision :  
   property Precision:Byte read Get_Precision write Set_Precision;
    // NumericScale :  
   property NumericScale:Byte read Get_NumericScale write Set_NumericScale;
    // OriginalValue :  
   property OriginalValue:OleVariant read Get_OriginalValue;
    // UnderlyingValue :  
   property UnderlyingValue:OleVariant read Get_UnderlyingValue;
    // DataFormat :  
   property DataFormat:IUnknown read Get_DataFormat write Set_DataFormat;
  end;


// Field20 : 

 Field20Disp = dispinterface
   ['{0000054C-0000-0010-8000-00AA006D2EA4}']
    // AppendChunk :  
   procedure AppendChunk(Data:OleVariant);dispid 1107;
    // GetChunk :  
   function GetChunk(Length:Integer):OleVariant;dispid 1108;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActualSize :  
   property ActualSize:ADO_LONGPTR  readonly dispid 1109;
    // Attributes :  
   property Attributes:Integer dispid 1114;
    // DefinedSize :  
   property DefinedSize:ADO_LONGPTR dispid 1103;
    // Name :  
   property Name:WideString  readonly dispid 1100;
    // Type :  
   property Type_:DataTypeEnum dispid 1102;
    // Value :  
   property Value:OleVariant dispid 0;
    // Precision :  
   property Precision:Byte dispid 1112;
    // NumericScale :  
   property NumericScale:Byte dispid 1113;
    // OriginalValue :  
   property OriginalValue:OleVariant  readonly dispid 1104;
    // UnderlyingValue :  
   property UnderlyingValue:OleVariant  readonly dispid 1105;
    // DataFormat :  
   property DataFormat:IUnknown dispid 1115;
  end;


// Field : 

 Field = interface(Field20)
   ['{00000569-0000-0010-8000-00AA006D2EA4}']
   function Get_Status : Integer; safecall;
    // Status :  
   property Status:Integer read Get_Status;
  end;

// Field : 

 FieldDisp = dispinterface
   ['{00000569-0000-0010-8000-00AA006D2EA4}']
    // AppendChunk :  
   procedure AppendChunk(Data:OleVariant);dispid 1107;
    // GetChunk :  
   function GetChunk(Length:Integer):OleVariant;dispid 1108;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActualSize :  
   property ActualSize:ADO_LONGPTR  readonly dispid 1109;
    // Attributes :  
   property Attributes:Integer dispid 1114;
    // DefinedSize :  
   property DefinedSize:ADO_LONGPTR dispid 1103;
    // Name :  
   property Name:WideString  readonly dispid 1100;
    // Type :  
   property Type_:DataTypeEnum dispid 1102;
    // Value :  
   property Value:OleVariant dispid 0;
    // Precision :  
   property Precision:Byte dispid 1112;
    // NumericScale :  
   property NumericScale:Byte dispid 1113;
    // OriginalValue :  
   property OriginalValue:OleVariant  readonly dispid 1104;
    // UnderlyingValue :  
   property UnderlyingValue:OleVariant  readonly dispid 1105;
    // DataFormat :  
   property DataFormat:IUnknown dispid 1115;
    // Status :  
   property Status:Integer  readonly dispid 1116;
  end;


// _Parameter : 

 _Parameter = interface(_ADO)
   ['{0000050C-0000-0010-8000-00AA006D2EA4}']
   function Get_Name : WideString; safecall;
   procedure Set_Name(const pbstr:WideString); safecall;
   function Get_Value : OleVariant; safecall;
   procedure Set_Value(const pvar:OleVariant); safecall;
   function Get_Type_ : DataTypeEnum; safecall;
   procedure Set_Type_(const psDataType:DataTypeEnum); safecall;
   procedure Set_Direction(const plParmDirection:ParameterDirectionEnum); safecall;
   function Get_Direction : ParameterDirectionEnum; safecall;
   procedure Set_Precision(const pbPrecision:Byte); safecall;
   function Get_Precision : Byte; safecall;
   procedure Set_NumericScale(const pbScale:Byte); safecall;
   function Get_NumericScale : Byte; safecall;
   procedure Set_Size(const pl:ADO_LONGPTR); safecall;
   function Get_Size : ADO_LONGPTR; safecall;
    // AppendChunk :  
   procedure AppendChunk(Val:OleVariant);safecall;
   function Get_Attributes : Integer; safecall;
   procedure Set_Attributes(const plParmAttribs:Integer); safecall;
    // Name :  
   property Name:WideString read Get_Name write Set_Name;
    // Value :  
   property Value:OleVariant read Get_Value write Set_Value;
    // Type :  
   property Type_:DataTypeEnum read Get_Type_ write Set_Type_;
    // Direction :  
   property Direction:ParameterDirectionEnum read Get_Direction write Set_Direction;
    // Precision :  
   property Precision:Byte read Get_Precision write Set_Precision;
    // NumericScale :  
   property NumericScale:Byte read Get_NumericScale write Set_NumericScale;
    // Size :  
   property Size:ADO_LONGPTR read Get_Size write Set_Size;
    // Attributes :  
   property Attributes:Integer read Get_Attributes write Set_Attributes;
  end;


// _Parameter : 

 _ParameterDisp = dispinterface
   ['{0000050C-0000-0010-8000-00AA006D2EA4}']
    // AppendChunk :  
   procedure AppendChunk(Val:OleVariant);dispid 7;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // Name :  
   property Name:WideString dispid 1;
    // Value :  
   property Value:OleVariant dispid 0;
    // Type :  
   property Type_:DataTypeEnum dispid 2;
    // Direction :  
   property Direction:ParameterDirectionEnum dispid 3;
    // Precision :  
   property Precision:Byte dispid 4;
    // NumericScale :  
   property NumericScale:Byte dispid 5;
    // Size :  
   property Size:ADO_LONGPTR dispid 6;
    // Attributes :  
   property Attributes:Integer dispid 8;
  end;


// Parameters : 

 Parameters = interface(_DynaCollection)
   ['{0000050D-0000-0010-8000-00AA006D2EA4}']
   function Get_Item(Index:OleVariant) : _Parameter; safecall;
    // Item :  
   property Item[Index:OleVariant]:_Parameter read Get_Item; default;
  end;


// Parameters : 

 ParametersDisp = dispinterface
   ['{0000050D-0000-0010-8000-00AA006D2EA4}']
    // _NewEnum :  
   function _NewEnum:IUnknown;dispid -4;
    // Refresh :  
   procedure Refresh;dispid 2;
    // Append :  
   procedure Append(Object_:IDispatch);dispid 1610809344;
    // Delete :  
   procedure Delete(Index:OleVariant);dispid 1610809345;
    // Count :  
   property Count:Integer  readonly dispid 1;
    // Item :  
   property Item[Index:OleVariant]:_Parameter  readonly dispid 0; default;
  end;


// Command25 : 

 Command25 = interface(Command15)
   ['{0000054E-0000-0010-8000-00AA006D2EA4}']
   function Get_State : Integer; safecall;
    // Cancel :  
   procedure Cancel;safecall;
    // State :  
   property State:Integer read Get_State;
  end;


// Command25 : 

 Command25Disp = dispinterface
   ['{0000054E-0000-0010-8000-00AA006D2EA4}']
    // Execute :  
   function Execute(out RecordsAffected:OleVariant;var Parameters:OleVariant;Options:Integer):_Recordset;dispid 5;
    // CreateParameter :  
   function CreateParameter(Name:WideString;Type_:DataTypeEnum;Direction:ParameterDirectionEnum;Size:ADO_LONGPTR;Value:OleVariant):_Parameter;dispid 6;
    // Cancel :  
   procedure Cancel;dispid 10;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActiveConnection :  
   property ActiveConnection:_Connection dispid 1;
    // CommandText :  
   property CommandText:WideString dispid 2;
    // CommandTimeout :  
   property CommandTimeout:Integer dispid 3;
    // Prepared :  
   property Prepared:WordBool dispid 4;
    // Parameters :  
   property Parameters:Parameters  readonly dispid 0;
    // CommandType :  
   property CommandType:CommandTypeEnum dispid 7;
    // Name :  
   property Name:WideString dispid 8;
    // State :  
   property State:Integer  readonly dispid 9;
  end;


// _Command : 

 _Command = interface(Command25)
   ['{B08400BD-F9D1-4D02-B856-71D5DBA123E9}']
   procedure Set_CommandStream(const pvStream:IUnknown); safecall;
   function Get_CommandStream : OleVariant; safecall;
   procedure Set_Dialect(const pbstrDialect:WideString); safecall;
   function Get_Dialect : WideString; safecall;
   procedure Set_NamedParameters(const pfNamedParameters:WordBool); safecall;
   function Get_NamedParameters : WordBool; safecall;
    // Dialect :  
   property Dialect:WideString read Get_Dialect write Set_Dialect;
    // NamedParameters :  
   property NamedParameters:WordBool read Get_NamedParameters write Set_NamedParameters;
  end;


// _Command : 

 _CommandDisp = dispinterface
   ['{B08400BD-F9D1-4D02-B856-71D5DBA123E9}']
    // Execute :  
   function Execute(out RecordsAffected:OleVariant;var Parameters:OleVariant;Options:Integer):_Recordset;dispid 5;
    // CreateParameter :  
   function CreateParameter(Name:WideString;Type_:DataTypeEnum;Direction:ParameterDirectionEnum;Size:ADO_LONGPTR;Value:OleVariant):_Parameter;dispid 6;
    // Cancel :  
   procedure Cancel;dispid 10;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActiveConnection :  
   property ActiveConnection:_Connection dispid 1;
    // CommandText :  
   property CommandText:WideString dispid 2;
    // CommandTimeout :  
   property CommandTimeout:Integer dispid 3;
    // Prepared :  
   property Prepared:WordBool dispid 4;
    // Parameters :  
   property Parameters:Parameters  readonly dispid 0;
    // CommandType :  
   property CommandType:CommandTypeEnum dispid 7;
    // Name :  
   property Name:WideString dispid 8;
    // State :  
   property State:Integer  readonly dispid 9;
    // CommandStream :  
   property CommandStream:IUnknown dispid 11;
    // Dialect :  
   property Dialect:WideString dispid 12;
    // NamedParameters :  
   property NamedParameters:WordBool dispid 13;
  end;


// ConnectionEventsVt : 

 ConnectionEventsVt = interface(IUnknown)
   ['{00000402-0000-0010-8000-00AA006D2EA4}']
    // InfoMessage :  
   function InfoMessage(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // BeginTransComplete :  
   function BeginTransComplete(TransactionLevel:Integer;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // CommitTransComplete :  
   function CommitTransComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // RollbackTransComplete :  
   function RollbackTransComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // WillExecute :  
   function WillExecute(var Source:WideString;var CursorType:CursorTypeEnum;var LockType:LockTypeEnum;var Options:Integer;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection):HRESULT;stdcall;
    // ExecuteComplete :  
   function ExecuteComplete(RecordsAffected:Integer;pError:Error;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection):HRESULT;stdcall;
    // WillConnect :  
   function WillConnect(var ConnectionString:WideString;var UserID:WideString;var Password:WideString;var Options:Integer;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // ConnectComplete :  
   function ConnectComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
    // Disconnect :  
   function Disconnect(var adStatus:EventStatusEnum;pConnection:_Connection):HRESULT;stdcall;
  end;


// RecordsetEventsVt : 

 RecordsetEventsVt = interface(IUnknown)
   ['{00000403-0000-0010-8000-00AA006D2EA4}']
    // WillChangeField :  
   function WillChangeField(cFields:Integer;Fields:OleVariant;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // FieldChangeComplete :  
   function FieldChangeComplete(cFields:Integer;Fields:OleVariant;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // WillChangeRecord :  
   function WillChangeRecord(adReason:EventReasonEnum;cRecords:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // RecordChangeComplete :  
   function RecordChangeComplete(adReason:EventReasonEnum;cRecords:Integer;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // WillChangeRecordset :  
   function WillChangeRecordset(adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // RecordsetChangeComplete :  
   function RecordsetChangeComplete(adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // WillMove :  
   function WillMove(adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // MoveComplete :  
   function MoveComplete(adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // EndOfRecordset :  
   function EndOfRecordset(var fMoreData:WordBool;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // FetchProgress :  
   function FetchProgress(Progress:Integer;MaxProgress:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
    // FetchComplete :  
   function FetchComplete(pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HRESULT;stdcall;
  end;


// ConnectionEvents : 

 ConnectionEvents = dispinterface
   ['{00000400-0000-0010-8000-00AA006D2EA4}']
    // InfoMessage :  
   function InfoMessage(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 0;
    // BeginTransComplete :  
   function BeginTransComplete(TransactionLevel:Integer;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 1;
    // CommitTransComplete :  
   function CommitTransComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 3;
    // RollbackTransComplete :  
   function RollbackTransComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 2;
    // WillExecute :  
   function WillExecute(var Source:WideString;var CursorType:CursorTypeEnum;var LockType:LockTypeEnum;var Options:Integer;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection):HResult;dispid 4;
    // ExecuteComplete :  
   function ExecuteComplete(RecordsAffected:Integer;pError:Error;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection):HResult;dispid 5;
    // WillConnect :  
   function WillConnect(var ConnectionString:WideString;var UserID:WideString;var Password:WideString;var Options:Integer;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 6;
    // ConnectComplete :  
   function ConnectComplete(pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 7;
    // Disconnect :  
   function Disconnect(var adStatus:EventStatusEnum;pConnection:_Connection):HResult;dispid 8;
  end;


// RecordsetEvents : 

 RecordsetEvents = dispinterface
   ['{00000266-0000-0010-8000-00AA006D2EA4}']
    // WillChangeField :  
   function WillChangeField(cFields:Integer;Fields:OleVariant;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 9;
    // FieldChangeComplete :  
   function FieldChangeComplete(cFields:Integer;Fields:OleVariant;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 10;
    // WillChangeRecord :  
   function WillChangeRecord(adReason:EventReasonEnum;cRecords:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 11;
    // RecordChangeComplete :  
   function RecordChangeComplete(adReason:EventReasonEnum;cRecords:Integer;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 12;
    // WillChangeRecordset :  
   function WillChangeRecordset(adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 13;
    // RecordsetChangeComplete :  
   function RecordsetChangeComplete(adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 14;
    // WillMove :  
   function WillMove(adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 15;
    // MoveComplete :  
   function MoveComplete(adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 16;
    // EndOfRecordset :  
   function EndOfRecordset(var fMoreData:WordBool;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 17;
    // FetchProgress :  
   function FetchProgress(Progress:Integer;MaxProgress:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 18;
    // FetchComplete :  
   function FetchComplete(pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset):HResult;dispid 19;
  end;


// ADOConnectionConstruction15 : 

 ADOConnectionConstruction15 = interface(IUnknown)
   ['{00000516-0000-0010-8000-00AA006D2EA4}']
   function Get_DSO : IUnknown; stdcall;
   function Get_Session : IUnknown; stdcall;
    // WrapDSOandSession :  
   function WrapDSOandSession(pDSO:IUnknown;pSession:IUnknown):HRESULT;stdcall;
    // DSO :  
   property DSO:IUnknown read Get_DSO;
    // Session :  
   property Session:IUnknown read Get_Session;
  end;


// ADOConnectionConstruction : 

 ADOConnectionConstruction = interface(ADOConnectionConstruction15)
   ['{00000551-0000-0010-8000-00AA006D2EA4}']
  end;


// _Record : 

 _Record = interface(_ADO)
   ['{00000562-0000-0010-8000-00AA006D2EA4}']
   function Get_ActiveConnection : OleVariant; safecall;
   procedure Set_ActiveConnection(const pvar:WideString); safecall;
   procedure Set_ActiveConnection_(const pvar:_Connection); safecall;
   function Get_State : ObjectStateEnum; safecall;
   function Get_Source : OleVariant; safecall;
   procedure Set_Source(const pvar:WideString); safecall;
   procedure Set_Source_(const pvar:IDispatch); safecall;
   function Get_Mode : ConnectModeEnum; safecall;
   procedure Set_Mode(const pMode:ConnectModeEnum); safecall;
   function Get_ParentURL : WideString; safecall;
    // MoveRecord :  
   function MoveRecord(Source:WideString;Destination:WideString;UserName:WideString;Password:WideString;Options:MoveRecordOptionsEnum;Async:WordBool):WideString;safecall;
    // CopyRecord :  
   function CopyRecord(Source:WideString;Destination:WideString;UserName:WideString;Password:WideString;Options:CopyRecordOptionsEnum;Async:WordBool):WideString;safecall;
    // DeleteRecord :  
   procedure DeleteRecord(Source:WideString;Async:WordBool);safecall;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);safecall;
    // Close :  
   procedure Close;safecall;
   function Get_Fields : Fields; safecall;
   function Get_RecordType : RecordTypeEnum; safecall;
    // GetChildren :  
   function GetChildren:_Recordset;safecall;
    // Cancel :  
   procedure Cancel;safecall;
    // State :  
   property State:ObjectStateEnum read Get_State;
    // Mode :  
   property Mode:ConnectModeEnum read Get_Mode write Set_Mode;
    // ParentURL :  
   property ParentURL:WideString read Get_ParentURL;
    // Fields :  
   property Fields:Fields read Get_Fields;
    // RecordType :  
   property RecordType:RecordTypeEnum read Get_RecordType;
  end;


// _Record : 

 _RecordDisp = dispinterface
   ['{00000562-0000-0010-8000-00AA006D2EA4}']
    // MoveRecord :  
   function MoveRecord(Source:WideString;Destination:WideString;UserName:WideString;Password:WideString;Options:MoveRecordOptionsEnum;Async:WordBool):WideString;dispid 6;
    // CopyRecord :  
   function CopyRecord(Source:WideString;Destination:WideString;UserName:WideString;Password:WideString;Options:CopyRecordOptionsEnum;Async:WordBool):WideString;dispid 7;
    // DeleteRecord :  
   procedure DeleteRecord(Source:WideString;Async:WordBool);dispid 8;
    // Open :  
   procedure Open(Source:OleVariant;ActiveConnection:OleVariant;Mode:ConnectModeEnum;CreateOptions:RecordCreateOptionsEnum;Options:RecordOpenOptionsEnum;UserName:WideString;Password:WideString);dispid 9;
    // Close :  
   procedure Close;dispid 10;
    // GetChildren :  
   function GetChildren:_Recordset;dispid 12;
    // Cancel :  
   procedure Cancel;dispid 13;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActiveConnection :  
   property ActiveConnection:OleVariant dispid 1;
    // State :  
   property State:ObjectStateEnum  readonly dispid 2;
    // Source :  
   property Source:OleVariant dispid 3;
    // Mode :  
   property Mode:ConnectModeEnum dispid 4;
    // ParentURL :  
   property ParentURL:WideString  readonly dispid 5;
    // Fields :  
   property Fields:Fields  readonly dispid 0;
    // RecordType :  
   property RecordType:RecordTypeEnum  readonly dispid 11;
  end;


// _Stream : 

 _Stream = interface(IDispatch)
   ['{00000565-0000-0010-8000-00AA006D2EA4}']
   function Get_Size : ADO_LONGPTR; safecall;
   function Get_EOS : WordBool; safecall;
   function Get_Position : ADO_LONGPTR; safecall;
   procedure Set_Position(const pPos:ADO_LONGPTR); safecall;
   function Get_Type_ : StreamTypeEnum; safecall;
   procedure Set_Type_(const ptype:StreamTypeEnum); safecall;
   function Get_LineSeparator : LineSeparatorEnum; safecall;
   procedure Set_LineSeparator(const pLS:LineSeparatorEnum); safecall;
   function Get_State : ObjectStateEnum; safecall;
   function Get_Mode : ConnectModeEnum; safecall;
   procedure Set_Mode(const pMode:ConnectModeEnum); safecall;
   function Get_Charset : WideString; safecall;
   procedure Set_Charset(const pbstrCharset:WideString); safecall;
    // Read_ :  
   function Read_(NumBytes:Integer):OleVariant;safecall;
    // Open :  
   procedure Open(Source:OleVariant;Mode:ConnectModeEnum;Options:StreamOpenOptionsEnum;UserName:WideString;Password:WideString);safecall;
    // Close :  
   procedure Close;safecall;
    // SkipLine :  
   procedure SkipLine;safecall;
    // Write_ :  
   procedure Write_(Buffer:OleVariant);safecall;
    // SetEOS :  
   procedure SetEOS;safecall;
    // CopyTo :  
   procedure CopyTo(DestStream:_Stream;CharNumber:ADO_LONGPTR);safecall;
    // Flush :  
   procedure Flush;safecall;
    // SaveToFile :  
   procedure SaveToFile(FileName:WideString;Options:SaveOptionsEnum);safecall;
    // LoadFromFile :  
   procedure LoadFromFile(FileName:WideString);safecall;
    // ReadText :  
   function ReadText(NumChars:Integer):WideString;safecall;
    // WriteText :  
   procedure WriteText(Data:WideString;Options:StreamWriteEnum);safecall;
    // Cancel :  
   procedure Cancel;safecall;
    // Size :  
   property Size:ADO_LONGPTR read Get_Size;
    // EOS :  
   property EOS:WordBool read Get_EOS;
    // Position :  
   property Position:ADO_LONGPTR read Get_Position write Set_Position;
    // Type :  
   property Type_:StreamTypeEnum read Get_Type_ write Set_Type_;
    // LineSeparator :  
   property LineSeparator:LineSeparatorEnum read Get_LineSeparator write Set_LineSeparator;
    // State :  
   property State:ObjectStateEnum read Get_State;
    // Mode :  
   property Mode:ConnectModeEnum read Get_Mode write Set_Mode;
    // Charset :  
   property Charset:WideString read Get_Charset write Set_Charset;
  end;


// _Stream : 

 _StreamDisp = dispinterface
   ['{00000565-0000-0010-8000-00AA006D2EA4}']
    // Read_ :  
   function Read_(NumBytes:Integer):OleVariant;dispid 9;
    // Open :  
   procedure Open(Source:OleVariant;Mode:ConnectModeEnum;Options:StreamOpenOptionsEnum;UserName:WideString;Password:WideString);dispid 10;
    // Close :  
   procedure Close;dispid 11;
    // SkipLine :  
   procedure SkipLine;dispid 12;
    // Write_ :  
   procedure Write_(Buffer:OleVariant);dispid 13;
    // SetEOS :  
   procedure SetEOS;dispid 14;
    // CopyTo :  
   procedure CopyTo(DestStream:_Stream;CharNumber:ADO_LONGPTR);dispid 15;
    // Flush :  
   procedure Flush;dispid 16;
    // SaveToFile :  
   procedure SaveToFile(FileName:WideString;Options:SaveOptionsEnum);dispid 17;
    // LoadFromFile :  
   procedure LoadFromFile(FileName:WideString);dispid 18;
    // ReadText :  
   function ReadText(NumChars:Integer):WideString;dispid 19;
    // WriteText :  
   procedure WriteText(Data:WideString;Options:StreamWriteEnum);dispid 20;
    // Cancel :  
   procedure Cancel;dispid 21;
    // Size :  
   property Size:ADO_LONGPTR  readonly dispid 1;
    // EOS :  
   property EOS:WordBool  readonly dispid 2;
    // Position :  
   property Position:ADO_LONGPTR dispid 3;
    // Type :  
   property Type_:StreamTypeEnum dispid 4;
    // LineSeparator :  
   property LineSeparator:LineSeparatorEnum dispid 5;
    // State :  
   property State:ObjectStateEnum  readonly dispid 6;
    // Mode :  
   property Mode:ConnectModeEnum dispid 7;
    // Charset :  
   property Charset:WideString dispid 8;
  end;


// ADORecordConstruction : 

 ADORecordConstruction = interface(IDispatch)
   ['{00000567-0000-0010-8000-00AA006D2EA4}']
   function Get_Row : IUnknown; safecall;
   procedure Set_Row(const ppRow:IUnknown); safecall;
   procedure Set_ParentRow(const Param1:IUnknown); safecall;
    // Row :  
   property Row:IUnknown read Get_Row write Set_Row;
    // ParentRow :  
   property ParentRow:IUnknown write Set_ParentRow;
  end;


// ADOStreamConstruction : 

 ADOStreamConstruction = interface(IDispatch)
   ['{00000568-0000-0010-8000-00AA006D2EA4}']
   function Get_Stream : IUnknown; safecall;
   procedure Set_Stream(const ppStm:IUnknown); safecall;
    // Stream :  
   property Stream:IUnknown read Get_Stream write Set_Stream;
  end;


// ADOCommandConstruction : 

 ADOCommandConstruction = interface(IUnknown)
   ['{00000517-0000-0010-8000-00AA006D2EA4}']
   function Get_OLEDBCommand : IUnknown; stdcall;
   procedure Set_OLEDBCommand(const ppOLEDBCommand:IUnknown); stdcall;
    // OLEDBCommand :  
   property OLEDBCommand:IUnknown read Get_OLEDBCommand write Set_OLEDBCommand;
  end;


// ADORecordsetConstruction : 

 ADORecordsetConstruction = interface(IDispatch)
   ['{00000283-0000-0010-8000-00AA006D2EA4}']
   function Get_Rowset : IUnknown; safecall;
   procedure Set_Rowset(const ppRowset:IUnknown); safecall;
   function Get_Chapter : ADO_LONGPTR; safecall;
   procedure Set_Chapter(const plChapter:ADO_LONGPTR); safecall;
   function Get_RowPosition : IUnknown; safecall;
   procedure Set_RowPosition(const ppRowPos:IUnknown); safecall;
    // Rowset :  
   property Rowset:IUnknown read Get_Rowset write Set_Rowset;
    // Chapter :  
   property Chapter:ADO_LONGPTR read Get_Chapter write Set_Chapter;
    // RowPosition :  
   property RowPosition:IUnknown read Get_RowPosition write Set_RowPosition;
  end;


// Field15 : 

 Field15 = interface(_ADO)
   ['{00000505-0000-0010-8000-00AA006D2EA4}']
   function Get_ActualSize : ADO_LONGPTR; safecall;
   function Get_Attributes : Integer; safecall;
   function Get_DefinedSize : ADO_LONGPTR; safecall;
   function Get_Name : WideString; safecall;
   function Get_Type_ : DataTypeEnum; safecall;
   function Get_Value : OleVariant; safecall;
   procedure Set_Value(const pvar:OleVariant); safecall;
   function Get_Precision : Byte; safecall;
   function Get_NumericScale : Byte; safecall;
    // AppendChunk :  
   procedure AppendChunk(Data:OleVariant);safecall;
    // GetChunk :  
   function GetChunk(Length:Integer):OleVariant;safecall;
   function Get_OriginalValue : OleVariant; safecall;
   function Get_UnderlyingValue : OleVariant; safecall;
    // ActualSize :  
   property ActualSize:ADO_LONGPTR read Get_ActualSize;
    // Attributes :  
   property Attributes:Integer read Get_Attributes;
    // DefinedSize :  
   property DefinedSize:ADO_LONGPTR read Get_DefinedSize;
    // Name :  
   property Name:WideString read Get_Name;
    // Type :  
   property Type_:DataTypeEnum read Get_Type_;
    // Value :  
   property Value:OleVariant read Get_Value write Set_Value;
    // Precision :  
   property Precision:Byte read Get_Precision;
    // NumericScale :  
   property NumericScale:Byte read Get_NumericScale;
    // OriginalValue :  
   property OriginalValue:OleVariant read Get_OriginalValue;
    // UnderlyingValue :  
   property UnderlyingValue:OleVariant read Get_UnderlyingValue;
  end;


// Field15 : 

 Field15Disp = dispinterface
   ['{00000505-0000-0010-8000-00AA006D2EA4}']
    // AppendChunk :  
   procedure AppendChunk(Data:OleVariant);dispid 1107;
    // GetChunk :  
   function GetChunk(Length:Integer):OleVariant;dispid 1108;
    // Properties :  
   property Properties:Properties  readonly dispid 500;
    // ActualSize :  
   property ActualSize:ADO_LONGPTR  readonly dispid 1109;
    // Attributes :  
   property Attributes:Integer  readonly dispid 1114;
    // DefinedSize :  
   property DefinedSize:ADO_LONGPTR  readonly dispid 1103;
    // Name :  
   property Name:WideString  readonly dispid 1100;
    // Type :  
   property Type_:DataTypeEnum  readonly dispid 1102;
    // Value :  
   property Value:OleVariant dispid 0;
    // Precision :  
   property Precision:Byte  readonly dispid 1112;
    // NumericScale :  
   property NumericScale:Byte  readonly dispid 1113;
    // OriginalValue :  
   property OriginalValue:OleVariant  readonly dispid 1104;
    // UnderlyingValue :  
   property UnderlyingValue:OleVariant  readonly dispid 1105;
  end;

//CoClasses
  TConnectionEventsInfoMessage = procedure(Sender: TObject;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsBeginTransComplete = procedure(Sender: TObject;TransactionLevel:Integer;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsCommitTransComplete = procedure(Sender: TObject;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsRollbackTransComplete = procedure(Sender: TObject;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsWillExecute = procedure(Sender: TObject;var Source:WideString;var CursorType:CursorTypeEnum;var LockType:LockTypeEnum;var Options:Integer;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection) of object;
  TConnectionEventsExecuteComplete = procedure(Sender: TObject;RecordsAffected:Integer;pError:Error;var adStatus:EventStatusEnum;pCommand:_Command;pRecordset:_Recordset;pConnection:_Connection) of object;
  TConnectionEventsWillConnect = procedure(Sender: TObject;var ConnectionString:WideString;var UserID:WideString;var Password:WideString;var Options:Integer;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsConnectComplete = procedure(Sender: TObject;pError:Error;var adStatus:EventStatusEnum;pConnection:_Connection) of object;
  TConnectionEventsDisconnect = procedure(Sender: TObject;var adStatus:EventStatusEnum;pConnection:_Connection) of object;


  CoConnection = Class
  Public
    Class Function Create: _Connection;
    Class Function CreateRemote(const MachineName: string): _Connection;
  end;

  TEvsConnection = Class(TEventSink)
  Private
    FOnInfoMessage:TConnectionEventsInfoMessage;
    FOnBeginTransComplete:TConnectionEventsBeginTransComplete;
    FOnCommitTransComplete:TConnectionEventsCommitTransComplete;
    FOnRollbackTransComplete:TConnectionEventsRollbackTransComplete;
    FOnWillExecute:TConnectionEventsWillExecute;
    FOnExecuteComplete:TConnectionEventsExecuteComplete;
    FOnWillConnect:TConnectionEventsWillConnect;
    FOnConnectComplete:TConnectionEventsConnectComplete;
    FOnDisconnect:TConnectionEventsDisconnect;

    fServer:_Connection;
    procedure EventSinkInvoke(Sender: TObject; DispID: Integer;
          const IID: TGUID; LocaleID: Integer; Flags: Word;
          Params: tagDISPPARAMS; VarResult, ExcepInfo, ArgErr: Pointer);
  Public
    constructor Create(TheOwner: TComponent); override;
    property ComServer:_Connection read fServer;
    property OnInfoMessage : TConnectionEventsInfoMessage read FOnInfoMessage write FOnInfoMessage;
    property OnBeginTransComplete : TConnectionEventsBeginTransComplete read FOnBeginTransComplete write FOnBeginTransComplete;
    property OnCommitTransComplete : TConnectionEventsCommitTransComplete read FOnCommitTransComplete write FOnCommitTransComplete;
    property OnRollbackTransComplete : TConnectionEventsRollbackTransComplete read FOnRollbackTransComplete write FOnRollbackTransComplete;
    property OnWillExecute : TConnectionEventsWillExecute read FOnWillExecute write FOnWillExecute;
    property OnExecuteComplete : TConnectionEventsExecuteComplete read FOnExecuteComplete write FOnExecuteComplete;
    property OnWillConnect : TConnectionEventsWillConnect read FOnWillConnect write FOnWillConnect;
    property OnConnectComplete : TConnectionEventsConnectComplete read FOnConnectComplete write FOnConnectComplete;
    property OnDisconnect : TConnectionEventsDisconnect read FOnDisconnect write FOnDisconnect;

  end;

  CoRecord = Class
  Public
    Class Function Create: _Record;
    Class Function CreateRemote(const MachineName: string): _Record;
  end;

  CoStream = Class
  Public
    Class Function Create: _Stream;
    Class Function CreateRemote(const MachineName: string): _Stream;
  end;

  CoCommand = Class
  Public
    Class Function Create: _Command;
    Class Function CreateRemote(const MachineName: string): _Command;
  end;

  TRecordsetEventsWillChangeField = procedure(Sender: TObject;cFields:Integer;Fields:OleVariant;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsFieldChangeComplete = procedure(Sender: TObject;cFields:Integer;Fields:OleVariant;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsWillChangeRecord = procedure(Sender: TObject;adReason:EventReasonEnum;cRecords:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsRecordChangeComplete = procedure(Sender: TObject;adReason:EventReasonEnum;cRecords:Integer;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsWillChangeRecordset = procedure(Sender: TObject;adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsRecordsetChangeComplete = procedure(Sender: TObject;adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsWillMove = procedure(Sender: TObject;adReason:EventReasonEnum;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsMoveComplete = procedure(Sender: TObject;adReason:EventReasonEnum;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsEndOfRecordset = procedure(Sender: TObject;var fMoreData:WordBool;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsFetchProgress = procedure(Sender: TObject;Progress:Integer;MaxProgress:Integer;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;
  TRecordsetEventsFetchComplete = procedure(Sender: TObject;pError:Error;var adStatus:EventStatusEnum;pRecordset:_Recordset) of object;


  CoRecordset = Class
  Public
    Class Function Create: _Recordset;
    Class Function CreateRemote(const MachineName: string): _Recordset;
  end;

  TEvsRecordset = Class(TEventSink)
  Private
    FOnWillChangeField:TRecordsetEventsWillChangeField;
    FOnFieldChangeComplete:TRecordsetEventsFieldChangeComplete;
    FOnWillChangeRecord:TRecordsetEventsWillChangeRecord;
    FOnRecordChangeComplete:TRecordsetEventsRecordChangeComplete;
    FOnWillChangeRecordset:TRecordsetEventsWillChangeRecordset;
    FOnRecordsetChangeComplete:TRecordsetEventsRecordsetChangeComplete;
    FOnWillMove:TRecordsetEventsWillMove;
    FOnMoveComplete:TRecordsetEventsMoveComplete;
    FOnEndOfRecordset:TRecordsetEventsEndOfRecordset;
    FOnFetchProgress:TRecordsetEventsFetchProgress;
    FOnFetchComplete:TRecordsetEventsFetchComplete;

    fServer:_Recordset;
    procedure EventSinkInvoke(Sender: TObject; DispID: Integer;
          const IID: TGUID; LocaleID: Integer; Flags: Word;
          Params: tagDISPPARAMS; VarResult, ExcepInfo, ArgErr: Pointer);
  Public
    constructor Create(TheOwner: TComponent); override;
    property ComServer:_Recordset read fServer;
    property OnWillChangeField : TRecordsetEventsWillChangeField read FOnWillChangeField write FOnWillChangeField;
    property OnFieldChangeComplete : TRecordsetEventsFieldChangeComplete read FOnFieldChangeComplete write FOnFieldChangeComplete;
    property OnWillChangeRecord : TRecordsetEventsWillChangeRecord read FOnWillChangeRecord write FOnWillChangeRecord;
    property OnRecordChangeComplete : TRecordsetEventsRecordChangeComplete read FOnRecordChangeComplete write FOnRecordChangeComplete;
    property OnWillChangeRecordset : TRecordsetEventsWillChangeRecordset read FOnWillChangeRecordset write FOnWillChangeRecordset;
    property OnRecordsetChangeComplete : TRecordsetEventsRecordsetChangeComplete read FOnRecordsetChangeComplete write FOnRecordsetChangeComplete;
    property OnWillMove : TRecordsetEventsWillMove read FOnWillMove write FOnWillMove;
    property OnMoveComplete : TRecordsetEventsMoveComplete read FOnMoveComplete write FOnMoveComplete;
    property OnEndOfRecordset : TRecordsetEventsEndOfRecordset read FOnEndOfRecordset write FOnEndOfRecordset;
    property OnFetchProgress : TRecordsetEventsFetchProgress read FOnFetchProgress write FOnFetchProgress;
    property OnFetchComplete : TRecordsetEventsFetchComplete read FOnFetchComplete write FOnFetchComplete;

  end;

  CoParameter = Class
  Public
    Class Function Create: _Parameter;
    Class Function CreateRemote(const MachineName: string): _Parameter;
  end;

implementation

uses comobj;

Class Function CoConnection.Create: _Connection;
begin
  Result := CreateComObject(CLASS_Connection) as _Connection;
end;

Class Function CoConnection.CreateRemote(const MachineName: string): _Connection;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Connection) as _Connection;
end;

constructor TEvsConnection.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  OnInvoke:=EventSinkInvoke;
  fServer:=CoConnection.Create;
  Connect(fServer,ConnectionEvents);
end;

procedure TEvsConnection.EventSinkInvoke(Sender: TObject; DispID: Integer;
  const IID: TGUID; LocaleID: Integer; Flags: Word; Params: tagDISPPARAMS;
  VarResult, ExcepInfo, ArgErr: Pointer);
begin
  case DispID of
    0: if assigned(OnInfoMessage) then
          OnInfoMessage(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    1: if assigned(OnBeginTransComplete) then
          OnBeginTransComplete(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    3: if assigned(OnCommitTransComplete) then
          OnCommitTransComplete(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    2: if assigned(OnRollbackTransComplete) then
          OnRollbackTransComplete(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    4: if assigned(OnWillExecute) then
          OnWillExecute(Self, WideString(Params.rgvarg[7].byRef^), CursorTypeEnum(Params.rgvarg[6].byRef^), LockTypeEnum(Params.rgvarg[5].byRef^), Params.rgvarg[4].plVal^, EventStatusEnum(Params.rgvarg[3].byRef^), OleVariant(Params.rgvarg[2]), OleVariant(Params.rgvarg[1]), OleVariant(Params.rgvarg[0]));
    5: if assigned(OnExecuteComplete) then
          OnExecuteComplete(Self, OleVariant(Params.rgvarg[5]), OleVariant(Params.rgvarg[4]), EventStatusEnum(Params.rgvarg[3].byRef^), OleVariant(Params.rgvarg[2]), OleVariant(Params.rgvarg[1]), OleVariant(Params.rgvarg[0]));
    6: if assigned(OnWillConnect) then
          OnWillConnect(Self, WideString(Params.rgvarg[5].byRef^), WideString(Params.rgvarg[4].byRef^), WideString(Params.rgvarg[3].byRef^), Params.rgvarg[2].plVal^, EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    7: if assigned(OnConnectComplete) then
          OnConnectComplete(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    8: if assigned(OnDisconnect) then
          OnDisconnect(Self, EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));

  end;
end;

Class Function CoRecord.Create: _Record;
begin
  Result := CreateComObject(CLASS_Record) as _Record;
end;

Class Function CoRecord.CreateRemote(const MachineName: string): _Record;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Record) as _Record;
end;

Class Function CoStream.Create: _Stream;
begin
  Result := CreateComObject(CLASS_Stream) as _Stream;
end;

Class Function CoStream.CreateRemote(const MachineName: string): _Stream;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Stream) as _Stream;
end;

Class Function CoCommand.Create: _Command;
begin
  Result := CreateComObject(CLASS_Command) as _Command;
end;

Class Function CoCommand.CreateRemote(const MachineName: string): _Command;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Command) as _Command;
end;

Class Function CoRecordset.Create: _Recordset;
begin
  Result := CreateComObject(CLASS_Recordset) as _Recordset;
end;

Class Function CoRecordset.CreateRemote(const MachineName: string): _Recordset;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Recordset) as _Recordset;
end;

constructor TEvsRecordset.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  OnInvoke:=EventSinkInvoke;
  fServer:=CoRecordset.Create;
  Connect(fServer,RecordsetEvents);
end;

procedure TEvsRecordset.EventSinkInvoke(Sender: TObject; DispID: Integer;
  const IID: TGUID; LocaleID: Integer; Flags: Word; Params: tagDISPPARAMS;
  VarResult, ExcepInfo, ArgErr: Pointer);
begin
  case DispID of
    9: if assigned(OnWillChangeField) then
          OnWillChangeField(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    10: if assigned(OnFieldChangeComplete) then
          OnFieldChangeComplete(Self, OleVariant(Params.rgvarg[4]), OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    11: if assigned(OnWillChangeRecord) then
          OnWillChangeRecord(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    12: if assigned(OnRecordChangeComplete) then
          OnRecordChangeComplete(Self, OleVariant(Params.rgvarg[4]), OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    13: if assigned(OnWillChangeRecordset) then
          OnWillChangeRecordset(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    14: if assigned(OnRecordsetChangeComplete) then
          OnRecordsetChangeComplete(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    15: if assigned(OnWillMove) then
          OnWillMove(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    16: if assigned(OnMoveComplete) then
          OnMoveComplete(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    17: if assigned(OnEndOfRecordset) then
          OnEndOfRecordset(Self, Params.rgvarg[2].pbool^, EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    18: if assigned(OnFetchProgress) then
          OnFetchProgress(Self, OleVariant(Params.rgvarg[3]), OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));
    19: if assigned(OnFetchComplete) then
          OnFetchComplete(Self, OleVariant(Params.rgvarg[2]), EventStatusEnum(Params.rgvarg[1].byRef^), OleVariant(Params.rgvarg[0]));

  end;
end;

Class Function CoParameter.Create: _Parameter;
begin
  Result := CreateComObject(CLASS_Parameter) as _Parameter;
end;

Class Function CoParameter.CreateRemote(const MachineName: string): _Parameter;
begin
  Result := CreateRemoteComObject(MachineName,CLASS_Parameter) as _Parameter;
end;

end.
