# three_sheet_

## Generalタブ (Sheet Data)



### General

| 表示項目名 | 使用部品 | 固定<br>/可変 | 表示<br>表示（入力）内容 | 表示<br>入力桁数 | 表示<br>配置 | 表示<br>デフォルト値 | Input/Output | チェック内容 | 備考 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Custumer data | GroupBox | 固定 | Custumer data |  | Left | - | - | - |  |
| Customer ID: | Static Text | 固定 | Customer ID: |  | Left | - | - | - |  |
| "Customer ID:"右TextBox | TextBox | 可変 | Don't Care | 20 | Left | 空 | CustomerApp.config | 空白不可、英数字のみ |  |
| Customer name: | Static Text | 固定 | Customer name: |  | Left | - |  | - |  |
| "Customer name:" 右TextBox | TextBox | 可変 | Don't Care | 40 | Left | 空 | CustomerApp.config | 空白不可、英数字のみ |  |
| Subsidiary: | Static Text | 固定 | Subsidiary: |  | Left | - |  | - |  |
| "Subsidiary:"右ComboBox | ComboBox | 可変 | Don't Care | - | Left | 空 | CustomerApp.config | 空白不可 | 現地法人Id：現地法人名 |
| Country: | Static Text | 固定 | Country: |  | Left | - |  | - |  |
| "Country:"右ComboBox | ComboBox | 可変 | Don't Care | - | Left | 空 | CustomerApp.config | 空白不可 | 国ID：国名 |
| Distributor ID: | Static Text | 固定 | Distributor ID: |  | Left | - |  | - |  |
| "Distributor ID:"右TextBox | TextBox | 可変 | Don't Care | 20 | Left | 空 | CustomerApp.config | 空白不可、英数字のみ |  |
| Distributor name: | Static Text | 固定 | Distributor name: |  | Left | - |  | - |  |
| "Distributor name:"右TextBox | TextBox | 可変 | Don't Care | 40 | Left | 空 | CustomerApp.config | 空白不可、英数字のみ |  |


### 空

| "Distributor name:"右TextBox | TextBox | 可変 | Don't Care | 40 | Left | 空 | 空白不可、英数字のみ |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Connecting instrument(s) List | GroupBox | 固定 | Custumer data |  | Left | - | - |
| 表ヘッダ1列目 | Static Text | 固定 | Status |  | center | - | - |
| 表ヘッダ2列目 | Static Text | 固定 | Instrument Type |  | center | - | - |
| 表ヘッダ3列目 | Static Text | 固定 | PS code |  | center | - | - |
| 表ヘッダ4列目 | Static Text | 固定 | Serial No. |  | center | - | - |
| DataGrid |  |  |  |  |  |  | 65以上登録状態でUpdateConfigするとエラー<br>データが存在してる場合は先頭データにフォーカスを当てておく |
| "Status"列 | ComboBox | 可変 | Enable/Disable | - | Left | 空 |  |
| "Instrument Type"列 | ComboBox | 可変 | Don't Care | - | Left | 空 | System設定の「Instrument list」タブで設定した情報を表示 |
| "PS code"列 | TextBox | 可変 | Don't Care | 無制限 | Left | 空 | 空白不可 |
| "Serial No."列 | TextBox | 可変 | Don't Care | 無制限 | Left | 空 | 空白不可 |


### 空

| "Serial No."列 | TextBox | 可変 | Don't Care | 無制限 | Left | 空 | 空白不可 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| "Delete" button | button | 固定 | Delete |  | center | - | 削除確認ダイアログ表示 |
| Instrument file: | Static Text | 固定 | Instrument file: |  | Left | - | - |
| "Instrument file:"下TextBox | TextBox | 可変 | Don't Care | 無制限 | Left | 空 | DataGrid選択時に、選択された機器の初期設定ファイルを表示<br>ファイル選択し確定した際にファイルパスを確認し、許容できない文字が含まれる場合はASCII文字チェックエラーダイアログを表示する。 |
| "Read from instrument file" button | button | 固定 | Read from instrument file |  | center | - | ボタン押下時の処理は外部仕様書参照。<br>ファイル存在チェックでFalseの場合は警告ダイアログ表示。<br>設定ファイル内が特定のフォーマットになっていない場合はListへの表示を行わない |
| BrowseFileSearchFolder | button | 固定 | .... |  | center | - | - |


![Generalタブ](images/three_sheet__Generalタブ_(Image)_image_33c45ff7.png)
