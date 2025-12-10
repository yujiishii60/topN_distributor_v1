# tools/cleanup-topn-split.ps1 ― 温惣菜／冷総菜ファイル削除ツール

## 概要
配布用 Excel 出力フォルダ内 (`data/output/split/`) にある  
**「温惣菜」「冷総菜」** を含むファイルをまとめて削除する PowerShell スクリプト。

---

## ファイル位置
```
C:\Users\14ugy\Projects\topN_distributor_v1\tools\cleanup-topn-split.ps1
```

---

## 使い方

### 1) DryRun（確認モード）
削除対象を確認するだけ。ファイルは削除されません。

```powershell
.\tools\cleanup-topn-split.ps1
```

### 2) 実削除モード
DryRun を明示的にオフにして実際に削除を行います。

```powershell
.\tools\cleanup-topn-split.ps1 -DryRun:$false
```

---

## 注意事項
- 対象はファイル名に **「温惣菜」または「冷総菜」** を含むもののみ。
- サブフォルダ内も再帰的に検索します。
- 削除後にフォルダが空になっても自動削除はしません。

---

## よくあるエラーと対処
| 症状 | 対処 |
|---|---|
| `ParameterArgumentTransformationError` | 別プロセスで `powershell -File ... -DryRun:$false` のように実行している可能性。現在のセッションで `.\tools\cleanup-topn-split.ps1 -DryRun:$false` と実行する。 |
| 文字化け | 出力に絵文字やUTF-8が混在する場合に発生。ASCIIメッセージ版のまま利用するのが安定。 |

---

## 更新履歴
| 日付 | 更新内容 |
|------|----------|
| 2025-11-09 | 初版作成（DryRun機能＋ASCII出力で安定動作） |
