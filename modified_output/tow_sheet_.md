# tow_sheet_

## XMLファイル自動生成 (Sheet Data)

&lt;TransferFileList&gt;タグ XMLファイル自動生成  
各機種の設定値シートから&lt;TransferFileList&gt;タグのXMLファイルを生成する  
### XMLファイル作成対象シート名

| **XMLファイル作成対象シート名**<br>転送設定初期値(RU-20)<br>転送設定初期値(CT-90) | **←①転送設定値が記載されたシート名を記載する(上詰めで！)** | **※注！　設定：AnalyzerTypeには対象機種の&lt;AnalyzerGroupID&gt;と同一の値を入れること！** |
| --- | --- | --- |
| 転送設定初期値(SP-10) |  | **その他は↓参照** |


### その他は↓参照

| 転送設定初期値(CT-90)<br>転送設定初期値(SP-10)<br>転送設定初期値(CF-60) | **※注！　設定：AnalyzerTypeには対象機種の&lt;AnalyzerGroupID&gt;と同一の値を入れること！**<br>**その他は↓参照** | IsEnabled | 「有効」or 「無効」 |
| --- | --- | --- | --- |
| 転送設定初期値(XEグループ) |  | FileName | 任意の文字列(必須) |
| 転送設定初期値(XTグループ) |  | FileType | 「Other」or「Measurement」or「Count」 |
| 転送設定初期値(XSグループ） |  | FileSearchFolder | 任意の文字列(必須) |
| 転送設定初期値(UX) |  | FileIncludeSearchDateSubfolders | 「OFF」or「ON」 |
| 転送設定初期値(RD-100i) |  | IsErrorEnable | 「OFF」or「ON」 |
| 転送設定初期値(CS-2x00i） |  | AnalyzerType | &lt;AnalyzerGroupID&gt;と同一の値 |
| 転送設定初期値(HISCL-2000i） |  | FileMask | 任意の文字列(必須) |
| 転送設定初期値(UFグループ） |  | ScaningType | 「Realtime」or「Timeperiod」 |
| 転送設定初期値(CS-5100i） |  | FileScanningInterval | 任意の数値 or 「N/A」 |
| 転送設定初期値(HISCL-800) |  | FileScanningIntervalMeasurementUnit | 「Day」or「Minute」or「N/A」 |
| 転送設定初期値(XN) |  | ProtcolType | 「Http」or「SMTP」 |
| 転送設定初期値(HISCL-5000） |  | FileEmailToAddress | 任意の文字列(指定なしの場合は「N/A」) |
| 転送設定初期値(RD-200)HTTP |  | ZipFile | 「OFF」or「ON」 |
| 転送設定初期値(CS-1600) |  | BackupDeleteFilesAfterTransfer | 「OFF」or「ON」 |
| 転送設定初期値(CS-2400) |  | BackupDeleteFilesElapsed | 「OFF」or「ON」 |
| 転送設定初期値(XN-L) |  | BackupFilesAfterTransfer | 「OFF」or「ON」 |
| 転送設定初期値(UF-Future) |  | BackupFolder | 任意の文字列(指定なしの場合は「N/A」) |
| 転送設定初期値(U-WAM） |  | BackupIncludeDateSubfolders | 「OFF」or「ON」 |
|  |  | BackupIncludeTimeStamp | 「OFF」or「ON」 |
|  |  | NumberOfDaysBackupFilesShouldBeKept | 任意の数値 |


←②ボタン押下  
## Instrument ID (Sheet Data)


### &lt;Instrument ID&gt;

| Group | Instrument Name | PS Code | Serial Number | Instrument ID | 備考 | 分析装置情報ファイル | 装置名 | PSコード | シリアル番号 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| XN | XN-20 | AE797961 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml | 検索タグ：&lt;Module_Info Position="0"&gt;装置名^PSコード^シリアル番号&lt;/Module_Info&gt; | 同左 | 同左 |
|  | XN-10 | AP795756 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml | 検索タグ：&lt;Module_Info Position="0"&gt;装置名^PSコード^シリアル番号&lt;/Module_Info&gt; | 同左 | 同左 |
| XN-L | XN-350 | AW618382 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-330 | CX851950 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-450 | BM392756 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-430 | CL032175 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-550 | BD634545 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-530 | BQ391107 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-150 | 未定 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | XN-130 | 未定 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
| XE | XE-5000 | 06375810 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-5000 JPN | C\Windows\XE-5000.ini | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  |  | 06376017 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-5000 EUR |  | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=98313314 | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  |  | 06376114 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-5000 AP |  | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=98313314 | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  |  | 06375917 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-5000 AMERICA |  | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=98313314 | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  | XE-2100 | 98313314 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100 JPN | C\Windows\XE-2100.ini | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[XE_Setting*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  |  | 98313411 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100 NA |  |  |  |  |
|  |  | 98313519 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100 EUR |  |  |  |  |
|  |  | 98313616 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100 ENG |  |  |  |  |
|  |  | 99337319 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100L NA |  |  |  |  |
|  |  | 99337211 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100L JPN |  |  |  |  |
|  |  | 99337416 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100L EUR |  |  |  |  |
|  |  | 99337513 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100L ENG |  |  |  |  |
|  |  | 03309313 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100D NA |  |  |  |  |
|  |  | 03309216 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100D JPN |  |  |  |  |
|  |  | 03309411 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100D EUR |  |  |  |  |
|  |  | 03309518 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100D ENG |  |  |  |  |
|  |  | 05323114 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100C NA |  |  |  |  |
|  |  | 05323017 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100C JPN |  |  |  |  |
|  |  | 05323319 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100DC NA |  |  |  |  |
|  |  | 05323211 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | XE-2100DC JPN |  |  |  |  |
| XT | XT-4000i | BY539823 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C\Windows\XT.ini | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  | XT-2000i | 01325318 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
|  | XT-1800i | 02305316 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
| XS | XS-1000i | 05342311 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\Sysmex\IPU\ORG_INI\IPU.ini | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  | XS-1000iC | BB046392 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
|  | XS-800i | 05347211 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
|  | XS-500i | CH642457 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
| SP | SP-10 | BR362571 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\Sysmex\SP-10\Send Folder\VER_*.XML | 検索タグ：&lt;Module_Info Position="0"&gt;装置名^PSコード^シリアル番号 | 同左 | 同左 |
| CF-60 | CF-60 | BN868986 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  |  |  |  |  |
| UF | UF-1000i | 05366719 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\sysmex\IPU\ORN_INI\IPU.ini | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
|  | UF-500i | AT487702 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\sysmex\IPU\ORN_INI\IPU.ini | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
| UF-Future | UF-3000 | AW402238 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | UF-4000 | BA212287 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | UF-5000 | BN344411 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
| UX | NO | AB541668 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\urinalysis\IPU\ORN_INI\IPU.ini | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 | 検索セクション：[IPU_System_Setting]<br>検索キー：PS_CODE=PSコード | 検索セクション：[Shx_Setting_*]<br>検索キー：InstrumentID_OriginalID=装置名^シリアル番号 |
| CS-1600 | CS-1600 | BQ203979 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | CS-1300 | CG705347 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
| CS-2400 | CS-2400 | AR780981 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | CS-2500 | BV981798 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | CS-5100 | BY990757 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 00-17以上 | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
| CS-2000 | CS-2000i | 06317410 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\CASettings\CA.System.Setting.xml | 検索タグ：&lt;Instrument_ID&gt;装置名^シリアル番号 | なし（ハードコーディングお願いします） | 検索タグ：&lt;Instrument_ID&gt;装置名^シリアル番号 |
|  | CS-2100i | 06372511 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\CASettings\CA.System.Setting.xml | 検索タグ：&lt;Instrument_ID&gt;装置名^シリアル番号 | なし（ハードコーディングお願いします） | 検索タグ：&lt;Instrument_ID&gt;装置名^シリアル番号 |
| RD-200 | RD-200 | BM372028 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | D:\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 |
| RD-100i | RD-100i | 05334814 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program Files\Sysmex\RmtMnt\SMTPSEND.ini | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 | 検索セクション：[RemoteMainte]<br>Subject =装置名^PSコード^シリアル番号 |
| HISCL-5000 | HISCL-5000 | AT911473 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内専用 | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml | 検索タグ：&lt;Module_Info&gt;装置名称^PSコード^シリアル番号&lt;/Module_Info&gt; | 同左 | 同左 |
|  |  | AF022051 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 海外用 | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml | 検索タグ：&lt;Module_Info&gt;装置名称^PSコード^シリアル番号&lt;/Module_Info&gt; | 同左 | 同左 |
|  | HISCL-800 | BW626143 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内専用 | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  |  | AE999369 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 海外用 | C:\Users\Public\Sysmex\Download\InstrumentInformation\InstrumentInformation_*.xml |  |  |  |
|  | HISCL-2000i | BK579523 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内専用（00-36以上） | 空文字 | ファイル内1行目：エラー通番,装置名称^PSコード^シリアル番号,日時,コード | 同左 | 同左 |
|  |  | CA759851 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 中国用(00-36以上） | 空文字 |  |  |  |
| CT90 | CT-90 | BD934079 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 海外はS/Nなし | C:\USERS\Public\Document\Sysmex\Download\InstrumentInformation\InstrumentInformation_[*].xml | 検索タグ：&lt;Module_Info Position="0"&gt;装置名^PSコード^シリアル番号&lt;/Module_Info&gt; | 同左 | 同左 |
|  | ST-40 | CJ256159 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | BT-40 | BT080317 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | CV-50 | AQ200900 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | CV-60 | CE944580 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | ST-42 | CV979713 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | ST-41 | AX575334 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
|  | TU-40 | AN617358 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No | 国内では使用しない |  |  |  |  |
| RU | RU-20 | CY662767 | 英数字5桁 | Instrument　Name^<br>PS Code^<br>Serial No |  | C:\Program files\sysmex\SNCS\SendFolder\RemoteMaintenace\EventHistory_*.XML | 検索タグ：&lt;Module_Info&gt;&lt;Instrument_ID&gt;装置名^PSコード^シリアル番号&lt;/Instrument_ID&gt;&lt;/Module_Info&gt; | 同左 | 同左 |


XN-9000の場合  
・ Other Regin  
CT-90のみ、設定する。本設定でST、BT、CV、TUのデータが送信されるように初期設定を実施する。  
ST、BT、CV、TUは設定する必要なし。（設定した場合、誤動作するので注意）  
なお、日本以外は、CT-90のシリアル番号は、BT-40のシリアル番号の先頭をAにして登録すること。  
SupportManagerインストール後のbatファイルについて  
市場では、必要なソフト、設定を実行後に、MECを導入して実行できるファイルを確定する。  
このため、batファイルを、SupportManagerインストール直後に実行するbatファイルと  
MEC導入後に実行するbatファイルに分けてもらう。  
## Sheet1 (Sheet Data)
