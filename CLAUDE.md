# hanamaru

日本株のテクニカル分析CLI(C#)。使い方・算出方法は `README.md` を参照。

## フォルダ構成

```
hanamaru/
├─ CLAUDE.md
├─ README.md            取扱説明書
├─ .gitignore           analysis/ と *.zip を除外
├─ src/
│  ├─ StockTool.csproj  ExcelDataReader 3.9.0 のみ参照
│  ├─ Program.cs        全ロジック(ta / screen / Yahoo取得 / 指標 / AIチェック)
│  ├─ bin/Release/      ビルド出力(git管理外)
│  └─ obj/              中間ファイル(git管理外)
└─ analysis/<日付>/     taの実行結果(git管理外)。カレントディレクトリ直下に作られる
   (screenのHTMLは hanamaru.exe と同じフォルダの analysis/yyyy-MM-dd.html)
```

## ビルド環境

- .NET SDK 10.0(`net10.0`、`win-x64`、Windows 専用)
- 単一ファイル・自己完結(`PublishSingleFile` / `SelfContained`)、出力名 `hanamaru.exe`
- ビルド: `cd src && dotnet publish -c Release`
- 出力先: `src/bin/Release/net10.0/win-x64/publish/hanamaru.exe`
- プログラムを修正したら、`dotnet build` で止めず必ず `dotnet publish -c Release` まで実行し、上記の `hanamaru.exe` を最新にしてから完了とする
- 警告 CA1416(`Registry.GetValue`)は Windows 専用なので無視してよい
- テストプロジェクトはない。動作確認は `hanamaru ta 7203` や `hanamaru screen 20` で行う

## 実行時の注意

- `ANTHROPIC_API_KEY` があるとスクリーニングの週足・日足◎銘柄をAIが採点(モデルは `Program.cs` の `Ai.Model`、現在 `claude-opus-5-5`)
- `STOCKS_PAR` で同時取得数を変更(既定24)
- 指標・判定・AIの仕様を変えたら `README.md` も更新する
- コミットメッセージ・返答は日本語で書く
- `POSITIONS_URL` / `POSITIONS_WRITE_URL`(ポジションシートの損切チェック `stops`)は Apps Script 方式。Sheets API(サービスアカウント)には切り替えない方針
