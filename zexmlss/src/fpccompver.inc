//For Delphi>=2009 compatibility in TObject added methods: ToString, GetHash, UnitName, Equals
{$IF FPC_FULLVERSION >= 20402} //FPC 2.4.2
  {$DEFINE Z_FPC_USE_TOSTRING}
{$IFEND}

//At leest in FPC 2.5.1 DecimalSeparator is deprecated, need use FormatSettings.DecimalSeparator
{$IF FPC_FULLVERSION >= 20402} //FPC 2.5.1
  {$DEFINE Z_USE_FORMAT_SETTINGS}
{$IFEND}
