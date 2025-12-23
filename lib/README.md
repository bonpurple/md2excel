# lib/ について（依存 jar の配置）

このプロジェクトは Maven/Gradle を使用しません。  
Apache POI 等の jar は **Git に含めない**ため、各自で取得して `lib/` に配置し、Eclipse の Build Path に追加してください。

---

## 入手先（どこから取得するか）

- Apache POI の公式配布物（Binary）から取得してください  
  - POI 本体 jar と、同梱されている依存 jar を利用します

> メモ: POI の配布物には `LICENSE` / `NOTICE` 等が含まれます。  
> 社内配布や再配布を行う場合は、それらの取り扱いに注意してください。

---

## 配置する jar（最低限）

`lib/` に以下の jar を配置してください（バージョンは揃えてください）。

- `poi-5.0.0.jar`
- `poi-ooxml-5.0.0.jar`
- `poi-ooxml-full-5.0.0.jar`
- `xmlbeans-4.0.0.jar`
- `commons-compress-1.20.jar`
- `commons-collections4-4.4.jar`

> 補足:
> - `.xlsx` を扱うため `poi-ooxml` / `poi-ooxml-full` が必要です。
> - jar が不足すると実行時に `ClassNotFoundException` が発生します。

---

## Eclipse 設定手順

1. プロジェクト直下に `lib/` フォルダを作成（または既存を使用）
2. 上記 jar を `lib/` に配置
3. Eclipse で `lib/` 内の jar を選択 → 右クリック → **Build Path** → **Add to Build Path**

---

## Git 管理について

- `lib/*.jar` は `.gitignore` 対象です（依存 jar はリポジトリにコミットしません）。
