# one_sheet_

## 転送設定初期値(UF-Future) (Sheet Data)

対象分析装置  
UF-Future  
### OnlineQC

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | OnlineQC | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\OnlineQC | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | ON | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | OnlineQC_*.txt | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### StartupReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | StartupReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | StartupReport_*.xml | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### ShutdownReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | ShutdownReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | ShutdownReport_*.xml | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### EventHistoryReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | EventHistoryReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | EventHistoryReport_*.csv | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### ErrorReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | ErrorReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RealtimeLog | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | ON | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | ErrorReport_*.xml | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### UserSettingBackup

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | UserSettingBackup | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | UserSettingBackup_*.ini | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### ServiceSettingBackup

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | ServiceSettingBackup | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | ServiceSettingBackup_*.ini | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### LogFile

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | LogFile | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RealtimeLog | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | LogFile_*.cab | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### SensitivityReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | SensitivityReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | SensitivityReport_*.* | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### OpticalReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | OpticalReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | Optical*Report_*.* | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### CalibrationReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | CalibrationReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | CalibrationReport_*.* | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### ConductivitySensorReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | ConductivitySensorReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | ConductivitySensorReport_*.* | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |


### AnalysisResultsReport

| 名前 | 初期値 | 設定値 |
| --- | --- | --- |
| TransferFileList |  |  |
| IsEnabled | 有効 | 転送設定が有効かどうか。無効な場合は転送対象外。 |
| FileName | AnalysisResultsReport | ファイル名称。自由入力 |
| FileType | Other | ファイルタイプ。Count,Measurement,Other |
| FileSearchFolder | D:\Sysmex\Upload\RemoteMaintenance | ファイルの検索フォルダ。 |
| FileIncludeSearchDateSubfolders | OFF | サブディレクトを検索するかどうか |
| IsErrorEnable | OFF | 転送失敗時にエラー画面を表示するかどうかの設定 |
| AnalyzerType | UF-Future | 分析装置タイプ |
| FileMask | AnalysisResultsReport_*.cab | ファイルマスクの文字列 |
| ScaningType | Realtime | スキャンタイプ。時刻orリアルタイム |
| FileScanningInterval | 5 | 監視間隔時間。スキャンタイプが時刻の時のみ有効。 |
| FileScanningIntervalMeasurementUnit | Minute | 監視間隔時間の単位。スキャンタイプが時刻の時のみ有効。 |
| ProtcolType | Http | プロトコルタイプ。Http or SMTP |
| FileEmailToAddress | N/A | 転送先メールアドレス。;で複数登録可能。プロトコルタイプがSMTPの場合のみ有効。 |
| ZipFile | OFF | ファイルZIP圧縮するかどうか |
| BackupDeleteFilesAfterTransfer | OFF | 転送後にファイルを削除するかどうか |
| BackupDeleteFilesElapsed | OFF | 古い転送ファイルと管理ファイルを削除するかどうか |
| BackupFilesAfterTransfer | OFF | 転送ファイルをバックアップするかどうか |
| BackupFolder | N/A | バックアップ先のフォルダ |
| BackupIncludeDateSubfolders | ON | バックアップフォルダにサブフォルダを作るかどうか |
| BackupIncludeTimeStamp | OFF | バックアップファイルにタイムスタンプをつけるかどうか |
| NumberOfDaysBackupFilesShouldBeKept | 0 | バックアップの日数 |

