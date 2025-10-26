# VisualExcelTaskForKanban

Excel のタスク一覧を読み込み、ドラッグ&ドロップ可能なカンバンボードとして可視化するデスクトップアプリケーションです。PyWebView を介してローカル HTML を表示し、Excel ファイルの編集・保存までを Python バックエンドで担います。

## ディレクトリ構成

```
VisualExcelTaskForKanban/
├─ backend/              # Python バックエンド
├─ data/                 # サンプルの Excel データなど
├─ frontend/
│  ├─ pages/             # HTML エントリーポイント
│  └─ scripts/           # フロントエンドの JavaScript
├─ tools/                # 実行補助スクリプト
├─ LICENSE
└─ README.md
```

## 主な機能

### カンバンボード
- Excel (`data/task.xlsx`) の各行をカードとして表示し、ステータス列ごとに整理。
- ステータス列上でのドラッグ&ドロップによる進捗更新と即時反映。
- タスク追加・編集・削除モーダル。
- 担当者／キーワード／ステータス／期限（前後・期間）でのフィルタリング機能。
- 期限切れ・期限間近タスクのハイライトと件数サマリー。
- 優先度列の昇順ソート（数値・文字列どちらも対応）。
- Excel のデータ検証（プルダウン）設定を読み取り、ステータス候補を自動で同期。
- 入力規則（候補値）の編集 UI。保存時は Excel 側のデータ検証も更新。

### タイムラインビュー (`frontend/pages/timeline.html`)
- 担当者単位でタスクを縦軸に、日付を横軸に並べた簡易ガントチャート風表示。
- 表示期間の指定（任意の日付範囲、今週・来週などのクイック選択）。
- ステータス／担当者／キーワードによるフィルタリングと集計。
- カンバンビューへ戻るショートカットボタン。

### Excel 連携
- 必須列は `No / ステータス / タスク / 担当者 / 優先度 / 期限 / 備考`。
- Excel に該当列が不足している場合は自動で追加。
- 優先度列が未追加の場合のマイグレーション手順:
  1. カンバンで利用している Excel ファイル（既定では `data/task.xlsx`）を Excel 等で開く。
  2. `担当者` 列の右隣に新しい列を挿入し、列名を **優先度** に変更。
  3. 既存タスクの優先度を必要に応じて入力（数値・文字列どちらでも可）。
  4. ファイルを保存。

## 動作環境

| 項目 | 内容 |
| ---- | ---- |
| OS | Windows / macOS / Linux （PyWebView が動作する環境） |
| Python | 3.10 以上を推奨（`str \| None` などのシンタックスを使用） |
| Excel | `data/task.xlsx` を含む任意の .xlsx ファイル |

### 必要な Python パッケージ

```
pip install pandas openpyxl pywebview watchdog
```

※ Windows で WebView2 ランタイムが未導入の場合は [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/#download-section) のインストールが必要になることがあります。

## 実行方法

1. 上記パッケージをインストールした Python 仮想環境を用意します。
2. `data/task.xlsx`（または自分の Excel ファイルを `data/` 配下へコピーしたもの）を用意します。
3. 次のコマンドでバックエンドを起動します。

```bash
python backend/backend.py --excel ./data/task.xlsx --html ./frontend/pages/index.html
```

PyWebView ウィンドウが開き、カンバンボードが表示されます。ツールバーの「保存」で Excel へ書き戻し、「再読込」で Excel から最新状態を再読み込みします。

### コマンドライン引数

`backend.py` では以下のオプションが利用できます。

| 引数 | 既定値 | 説明 |
| ---- | ------ | ---- |
| `--excel` | `./data/task.xlsx` | 読み込む Excel ファイルのパス |
| `--sheet` | なし | 読み込むシート名（未指定時は先頭シート） |
| `--html` | `./frontend/pages/index.html` | 表示するフロントエンド HTML のパス |
| `--title` | `タスク・ボード` | ウィンドウタイトル |
| `--width` | `1280` | ウィンドウ幅（ピクセル） |
| `--height` | `800` | ウィンドウ高さ（ピクセル） |
| `--debug` | `False` | PyWebView のデバッグモードを有効化（開発時向け） |
| `--no-watch` | `False` | Excel ファイルの変更監視を無効化 |
| `--watch-polling` | `False` | ファイル監視に PollingObserver を利用（ネットワークドライブ向け） |
| `--watch-interval` | `1.0` | PollingObserver 利用時のポーリング間隔（秒） |
| `--watch-debounce` | `2.0` | アプリ自身の保存直後に発生するイベントを無視する猶予時間（秒） |

### Windows 用バッチ

Windows で Miniconda を利用している場合は、`tools/VisualExcelTaskForKanban.bat` を編集して利用することで、仮想環境の有効化とアプリ起動をまとめて行えます。

### Excel ファイルの自動監視

- バックエンドは [watchdog](https://pypi.org/project/watchdog/) を利用して Excel ファイルの変更を常時監視し、保存を検知すると自動で `load_excel()` を実行します。
- 監視で取得した最新データは PyWebView 経由でフロントエンドへプッシュされ、手動の「再読込」操作なしでボードが更新されます。
- 監視が不要な場合は `--no-watch` を指定してください。ネットワークドライブなどの環境では `--watch-polling`（必要に応じて `--watch-interval`）でポーリング監視へ切り替えられます。
- アプリの保存直後に発生する監視イベントは `--watch-debounce` で指定した秒数だけ無視されるため、無限ループで再読込されることはありません。
- 動作確認はアプリ起動中に Excel を外部で編集・保存し、数秒後にボードへ自動反映されることを確認してください。

## ライセンス

このリポジトリは [MIT License](./LICENSE) の下で公開されています。
