# ExcelKanban_Lite

Excel のタスク一覧を読み込み、ドラッグ&ドロップ可能なカンバンボードとして可視化するデスクトップアプリケーションです。PyWebView を介してローカル HTML を表示し、Excel ファイルの読み込み・編集・保存、データ検証（入力規則）の同期、ファイル変更の自動反映までを Python バックエンドで担います。タスクはカンバン／リスト／タイムラインの 3 画面で閲覧・編集でき、いずれの画面からでも Excel 側の候補値や担当者別サマリーを共通で利用できます。

## ディレクトリ構成

```
ExcelKanban_Lite/
├─ backend/              # Python バックエンド
├─ data/                 # サンプルの Excel データなど
├─ frontend/
│  ├─ pages/             # HTML エントリーポイント（カンバン／リスト／タイムライン）
│  ├─ scripts/           # フロントエンドの JavaScript
│  └─ styles/            # 各画面のスタイルシート
├─ tools/                # 実行補助スクリプト
├─ LICENSE
└─ README.md
```

## 主な画面と機能

### カンバンボード (`frontend/pages/index.html`)
- Excel (`data/task.xlsx`) の各行をカードとして表示し、ステータス列ごとに整理。
- カードのドラッグ&ドロップによるステータス移動・優先度更新や、モーダルからの追加／編集／削除を即時反映。【F:frontend/pages/index.html†L20-L121】【F:frontend/scripts/index.js†L1-L220】
- 担当者／大分類・中分類／複数ステータス／キーワード／期限（範囲・以前・以後）でのフィルタリングと一括解除ボタンを搭載。【F:frontend/pages/index.html†L29-L74】【F:frontend/scripts/index.js†L402-L563】
- 期限切れ・期限間近タスクの件数バッジとトースト通知をツールバーに表示し、カードにもハイライトを付与。【F:frontend/pages/index.html†L12-L26】【F:frontend/scripts/common.js†L326-L460】
- 担当者別のカード件数／ステータス内訳／期限警告をまとめた「担当者別サマリー」を表示し、担当者フィルターと連動。【F:frontend/pages/index.html†L76-L96】【F:frontend/scripts/common.js†L539-L676】
- Excel のデータ検証（プルダウン）設定を読み取り、ステータスや優先度などの候補を同期。入力規則編集モーダルから候補値を更新すると Excel 側の検証ルールも上書きされます。【F:frontend/pages/index.html†L122-L171】【F:frontend/scripts/common.js†L218-L323】【F:backend/backend.py†L214-L335】

### リストビュー (`frontend/pages/list.html`)
- すべてのタスクをテーブル形式で俯瞰し、列ヘッダーのクリックで複数列ソート（昇順／降順／リセット）を切り替え可能。【F:frontend/scripts/list.js†L947-L1030】
- 大分類／中分類ごとのグルーピング行と件数表示で階層構造を把握しやすく、ダブルクリックで編集モーダルを開けます。【F:frontend/scripts/list.js†L310-L356】【F:frontend/scripts/list.js†L880-L938】
- 列幅のドラッグ調整とローカルストレージへの保存で、自分好みのレイアウトを維持できます。【F:frontend/scripts/list.js†L200-L344】【F:frontend/scripts/list.js†L802-L833】
- 期限バッジや優先度ピルによる視認性向上、担当者別サマリーやフィルター UI はカンバンと共通でリアルタイムに同期。【F:frontend/pages/list.html†L14-L115】【F:frontend/scripts/list.js†L1-L199】【F:frontend/scripts/common.js†L539-L676】
- 横スクロールを補助する専用スクロールバーやキーボード操作に配慮したアクセシビリティ改善を実装。【F:frontend/pages/list.html†L117-L150】【F:frontend/scripts/list.js†L148-L233】【F:frontend/styles/list.css†L1-L120】

### タイムラインビュー (`frontend/pages/timeline.html`)
- 担当者単位でタスクを縦軸に、日付を横軸に並べた簡易ガントチャート風表示。表示期間外や期限未設定のタスクは「バックログ」欄に整理。【F:frontend/pages/timeline.html†L13-L74】【F:frontend/pages/timeline.html†L76-L96】
- 表示期間は自由入力に加え、「前週」「今週」「翌週」のワンクリック切り替えや開始・終了日の自動補完に対応。【F:frontend/pages/timeline.html†L32-L57】【F:frontend/scripts/timeline.js†L99-L186】
- 担当者フィルター、ステータス凡例、稼働サマリーで担当者ごとの負荷状況を可視化。ダブルクリックでタスク編集モーダルを呼び出せます。【F:frontend/pages/timeline.html†L30-L92】【F:frontend/scripts/timeline.js†L187-L388】
- カンバン／リストと同じ入力規則・優先度候補を共有し、保存後はすべての画面へ即座に反映されます。【F:frontend/scripts/timeline.js†L1-L154】【F:frontend/scripts/common.js†L218-L323】

### Excel 連携とリアルタイム更新
- 標準列は `No / ステータス / 大分類 / 中分類 / タスク / 担当者 / 優先度 / 期限 / 備考`。不足列は初回読込時に自動追加されます。【F:backend/backend.py†L188-L276】
- 保存時に Excel のデータ検証を上書きし、ステータスやカテゴリ候補を同期。保存ごとにタイムスタンプ付きバックアップ（`.bak_YYYYMMDD_HHMMSS.xlsx`）も生成します。【F:backend/backend.py†L357-L520】
- watchdog によるファイル監視で、Excel を外部で更新するとアプリへ自動プッシュ（ポーリング監視モードも選択可能）。【F:backend/backend.py†L522-L688】
- HTML をブラウザで直接開くとモックデータで UI を確認でき、PyWebView 非依存でデザイン検証が可能です。【F:frontend/scripts/common.js†L24-L167】

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

### 開発者向けヒント

- HTML をブラウザで直接開くとモックデータが表示され、PyWebView なしでも UI の確認ができます。
- Excel の変更監視が不要な場合は `--no-watch` を指定してください。ネットワークドライブなどは `--watch-polling` とポーリング間隔オプションで調整できます。

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

Windows で Miniconda を利用している場合は、`tools/ExcelKanban_Lite.bat` を編集して利用することで、仮想環境の有効化とアプリ起動をまとめて行えます。

### Excel ファイルの自動監視

- バックエンドは [watchdog](https://pypi.org/project/watchdog/) を利用して Excel ファイルの変更を常時監視し、保存を検知すると自動で `load_excel()` を実行します。
- 監視で取得した最新データは PyWebView 経由でフロントエンドへプッシュされ、手動の「再読込」操作なしでボードが更新されます。
- 監視が不要な場合は `--no-watch` を指定してください。ネットワークドライブなどの環境では `--watch-polling`（必要に応じて `--watch-interval`）でポーリング監視へ切り替えられます。
- アプリの保存直後に発生する監視イベントは `--watch-debounce` で指定した秒数だけ無視されるため、無限ループで再読込されることはありません。
- 動作確認はアプリ起動中に Excel を外部で編集・保存し、数秒後にボードへ自動反映されることを確認してください。

## おすすめポイント（このツールを使うメリット）

- **Excel 資産をそのまま活用して可視化**: 既存の Excel から不足列を自動補完しつつ、バックアップ生成と入力規則の上書きで安全に編集できます。【F:backend/backend.py†L188-L520】
- **担当者視点での負荷把握が簡単**: カンバン・リスト双方で担当者別サマリーを自動作成し、期限警告やステータス別内訳をワンクリックで確認可能です。【F:frontend/pages/index.html†L76-L96】【F:frontend/scripts/common.js†L539-L676】
- **場面に応じた 3 画面をワンタッチ切替**: カンバンで進捗管理、リストでレポート出力、タイムラインでガント風確認と、用途に合わせて同じデータを切り替えて使えます。【F:frontend/pages/index.html†L20-L26】【F:frontend/pages/list.html†L14-L26】【F:frontend/pages/timeline.html†L13-L47】
- **リアルタイム連携とチーム編集**: Excel ファイルの外部更新を自動取り込みし、入力規則編集も含めた変更が全画面に同時反映されるため、チームでの同時利用に向いています。【F:backend/backend.py†L522-L688】【F:frontend/scripts/index.js†L60-L109】
- **UI カスタマイズも手軽**: 列幅の保存・ソート条件・フィルター・バックログ表示など、利用シーンに合わせた表示調整がフロントエンドだけで完結します。【F:frontend/scripts/list.js†L200-L344】【F:frontend/scripts/timeline.js†L99-L388】

## ライセンス

このリポジトリは [MIT License](./LICENSE) の下で公開されています。
