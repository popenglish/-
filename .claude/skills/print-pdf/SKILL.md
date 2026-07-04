---
name: print-pdf
description: 通知表を全員分いっきにPDF化するスクリプトをクラス用に作って、貼り付け手順を案内するスキル。「印刷したい」「全員分のPDFを作りたい」と言われたら使う。
---

# /print-pdf — 通知表の全員分PDF作成

クラス名を聞いて、そのクラス用の「全員分PDF作成」スクリプトを作り、Apps Script への貼り付け方から実行までを1ステップずつ案内する。

## 前提

- スプレッドシート：「イエロークラス成績表 2026年度前期」ID: `1USHJ1JGM-trt8bCdkh9qwT9ksJo1hAl6tAOYlP3vpeM`
- 各クラスに `入力_クラス名` と `通知表_クラス名` のタブがある
- 通知表の名前セルは **B2**（B2を変えると全項目が自動で切り替わる仕組み）
- 仕組み：スクリプトが生徒名を1人ずつB2にセットして、その都度A4のPDFを保存 → 日付つきフォルダにまとまる → フォルダからまとめて印刷

## 手順

1. **クラス名とタブ名を確認する**：「入力タブと通知表タブの名前をコピーして貼ってください」と頼む（全角・半角の事故を防ぐため、**実際のタブ名をそのままスクリプトに入れる**）
2. 名前セルが B2 でよいか、通知表のスクリーンショットか一言で確認
3. 下のスクリプトの `★` の3か所をそのクラス用に書き換えて渡す
4. 貼り付け手順を案内する（下記）
5. 実行と印刷の案内

## スクリプト（★の3か所をクラスに合わせて書き換える）

```javascript
function 全員分PDF作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 入力 = ss.getSheetByName('入力_中一A');      // ★1 入力タブ名
  const テンプレ = ss.getSheetByName('通知表_中一A'); // ★2 通知表タブ名
  const 名前セル = 'B2';                              // ★3 名前のセル

  const folder = DriveApp.createFolder('通知表_中一A_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm'));
  const names = 入力.getRange('A2:A' + 入力.getLastRow()).getValues().flat().filter(String);
  const token = ScriptApp.getOAuthToken();
  const gid = テンプレ.getSheetId();
  const base = 'https://docs.google.com/spreadsheets/d/' + ss.getId()
    + '/export?format=pdf&gid=' + gid
    + '&size=A4&portrait=true&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=false';

  names.forEach(name => {
    テンプレ.getRange(名前セル).setValue(name);
    SpreadsheetApp.flush();
    Utilities.sleep(400);
    const res = UrlFetchApp.fetch(base, { headers: { Authorization: 'Bearer ' + token } });
    folder.createFile(res.getBlob().setName(name + '.pdf'));
  });
  SpreadsheetApp.getActive().toast(names.length + '人分のPDFを作成しました ✅', '完了', 6);
}
```

- 複数クラスをまとめてやりたい場合は、関数名を `全員分PDF作成_中一A` のようにクラスごとに分けて、同じファイルに並べてよい（フォルダ名も合わせて変える）

## 貼り付け手順（ユーザーへの案内。1ステップずつ）

1. スプレッドシートを開く → 上のメニュー「**拡張機能**」→「**Apps Script**」
2. 開いた画面に既にコードがあれば、**一番下の空いた行**に貼る（既存のコードは消さない）。`function` の途中に貼らないこと
3. 貼ったら **⌘S で保存**
4. 上の関数選択（▼）で `全員分PDF作成` を選んで「**実行**」
5. 初回は許可画面が出る：アカウントを選ぶ →「詳細」→「（安全でないページに）移動」→「許可」
   （自分で作ったスクリプトなので安全。Googleの決まり文句です、と安心させる）
6. 完了すると「◯人分のPDFを作成しました ✅」と出る → Googleドライブに日付つきフォルダができている
7. フォルダを開いて全選択 → ダウンロードするか、1つずつ開いて印刷（Macならフォルダごとダウンロード→プレビューでまとめて開いて⌘Pが速い）

## よくあるつまずき

- **エラー「null の getRange」**→ タブ名がスクリプトと実際で違う。タブ名をコピーして★1★2に貼り直す（全角・半角に注意）
- **PDFの中身が全員同じ**→ 名前セル（★3）の場所が違う。通知表で名前が出るセルの番地を確認
- **見切れる・2ページになる**→ 通知表の列幅・行数をA4に収まるよう調整するか、export URLに `&scale=4`（幅に合わせる）が入っているか確認
- 実行前に、通知表の数式（VLOOKUP）が正しく動いているか1人分で確認してもらう

## 返事のしかた

- 日本語で、1ステップずつ。「できた」を待ってから次へ
- スクリプトはユーザーのアカウントで動くので権限の心配はない、と伝える
