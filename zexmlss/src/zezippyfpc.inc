//Какого хрена дельфа зашла в {$IFDEF FPC}???
{$IF FPC_FULLVERSION < 20501} //FPC 2.5.1
  {$DEFINE USE_INTERNAL_SL}
{$IFEND}