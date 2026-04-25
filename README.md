# 🍛 カレー食品加工 在庫管理システム

Claude Code で作成する、Google スプレッドシート連携の在庫管理 Web アプリです。
スマホ・PC どちらからでも使えて、ボタン1つで入庫・出庫が記録できます。

## 機能

- 📦 **在庫マスタ管理** - 商品・カテゴリー・容器(単位)・最低在庫の管理
- ➕➖ **ワンタップ入庫・出庫** - スマホでも片手で操作可能
- 📜 **履歴記録** - いつ・誰が・何を・いくつ操作したか自動記録
- 🚨 **在庫アラート** - 最低在庫を下回った商品をハイライト表示
- 🏷️ **カテゴリー別表示** - 香辛料・野菜・肉類などタブで切り替え

## 構成

```
curry-inventory/
├── README.md                  ← このファイル
├── SETUP.md                   ← 詳細セットアップ手順
├── CLAUDE.md                  ← Claude Code 用コンテキスト
├── gas/
│   ├── Code.gs               ← Google Apps Script(API バックエンド)
│   └── appsscript.json       ← GAS 設定ファイル
├── web/
│   ├── index.html            ← メイン画面
│   ├── styles.css            ← スタイル
│   └── app.js                ← フロント JS
└── data/
    └── initial_data.md       ← 初期投入用サンプルデータ
```

## アーキテクチャ

```
[スマホ/PC ブラウザ]
       ↓ HTTPS (fetch API)
[Google Apps Script Web App]
       ↓ SpreadsheetApp
[Google スプレッドシート]
   - 在庫マスタ
   - 履歴
   - カテゴリー
   - 設定
```

サーバー不要・無料で動きます。

## はじめかた

`SETUP.md` を上から順番にやれば 30 分で動きます。

1. Google スプレッドシートを作る(雛形 SQL あり)
2. Apps Script を貼り付けてデプロイ
3. デプロイ URL を `web/app.js` に貼る
4. `web/index.html` をブラウザで開く

詳細は `SETUP.md` を見てください。
