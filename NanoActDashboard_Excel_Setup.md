# Excel VBA版セットアップ手順

## 作るもの
- `Dashboard` シート: KPI、フィルタ、一覧
- `Input` シート: 入力フォーム
- `Data` シート: 保存データ
- `Master` シート: プルダウン候補

## 1. 新しいExcelファイルを作る
1. Excelで新規ブックを作成
2. 名前を付けて保存
3. 形式は `Excel マクロ有効ブック (.xlsm)` にする

## 2. VBAを読み込む
1. `option + F11` でVBAエディタを開く
2. 左のプロジェクトを右クリック
3. `ファイルのインポート`
4. `/Users/yutasakurai/Documents/New project/NanoActDashboardVBA.bas` を選ぶ

## 3. マクロを実行してシートを自動作成する
1. VBAエディタで `NanoActDashboardVBA` モジュールを開く
2. `SetupWorkbook` の中にカーソルを置く
3. `F5` で実行
4. `Dashboard / Input / Data / Master` の4シートが作成される

## 4. 使い方
### 新規登録
1. `Input` シートを開く
2. 各項目を入力
3. `保存` ボタンを押す
4. `Data` シートに1件保存される
5. `Dashboard` シートに一覧とKPIが反映される

### 編集
1. `Dashboard` シートで対象行を選ぶ
2. `選択行を開く` を押す
3. `Input` シートに内容が読み込まれる
4. 修正して `保存`

### 削除
1. `Input` シートで対象データを読み込む
2. `削除` ボタンを押す

### GoogleMapを開く
1. `Input` シートの `GoogleMapリンク` にURLを入れる
2. `GoogleMapを開く` を押す

### CSV出力
1. `Dashboard` シートで `CSV出力` を押す
2. 保存先を選ぶ

## 5. フィルタ
- `Dashboard` シートの以下を入力すると一覧を絞り込めます
  - 種別フィルタ
  - ステージフィルタ
  - 温度感フィルタ
- 変更後に `更新` を押す

## 6. できること
- Saveボタンで保存
- 既存レコードの編集
- 削除
- KPI表示
- フィルタ
- CSV出力
- GoogleMapリンクを開く

## 7. 補足
- このExcel版はローカルファイル保存です
- 複数人で同時に編集・共有したい場合は、次段階で `Microsoft Lists` 連携に進めるのが適切です
- まずは Excel で業務フローを固める用途に向いています
