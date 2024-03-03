<!--
Copyright 2023 tsukuboshi

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
# gas-count-schedule-time

## 概要

Google Apps Scriptで、Google Calendarの予定をカウントしGoogle SpreadSheetに書き込んだ上で、予定のラベル(色)毎に工数合計を計算するスクリプトです。

## 前提条件

以下のソフトウェアの使用を前提とします。

- Node.js；v20以降
- TypeScript；v5以降

また、GASを実行するためのGoogleアカウントが必要です。  

## デプロイ手順

1. リポジトリをクローンし、ディレクトリを移動

```bash
git clone https://github.com/tsukuboshi/gas-count-schedule-time.git
```

2. npmパッケージをインストール

```bash
npm init -y
npm install
```

3. Googleアカウントの認証を実施

```bash
npx clasp login
```

4. 環境変数を設定(変数の値は適宜変更してください)

```bash
export CALENDER_ID_ARRAY="['xxxxxxxx','xxxxxxxx','xxxxxxxx']"
```

5. カレンダーに対応したラベル名を指定(必要ない場合は空文字を設定してください)

```bash
export DEFAULT_LABEL=""
export LAVENDER_LABEL=""
export SAGE_LABEL=""
export GRAPE_LABEL=""
export FLAMINGO_LABEL=""
export BANANA_LABEL=""
export MANDARIN_LABEL=""
export PEACOCK_LABEL=""
export GRAPHITE_LABEL=""
export BLUEBERRY_LABEL=""
export BASIL_LABEL=""
export TOMATO_LABEL=""
```

6. 初期化ファイルを作成

```bash

cat <<EOF > src/index.ts
import { main } from './example-module';

function handler(): void {
  main(
    ${CALENDER_ID_ARRAY},
    {
      デフォルト: '${DEFAULT_LABEL}',
      ラベンダー: '${LAVENDER_LABEL}',
      セージ: '${SAGE_LABEL}',
      ブドウ: '${GRAPE_LABEL}',
      フラミンゴ: '${FLAMINGO_LABEL}',
      バナナ: '${BANANA_LABEL}',
      ミカン: '${MANDARIN_LABEL}',
      ピーコック: '${PEACOCK_LABEL}',
      グラファイト: '${GRAPHITE_LABEL}',
      ブルーベリー: '${BLUEBERRY_LABEL}',
      バジル: '${BASIL_LABEL}',
      トマト: '${TOMATO_LABEL}',
    }
  );
}
EOF
```

7. GASスクリプトをデプロイ

```bash
npm run deploy
```

## 実行手順

1. 以下コマンドより、対象GASのプロジェクトを開いてください。

```bash
npx clasp open
```

2. プロジェクトの「エディター」より、`index.gs`にある`handler`関数を実行してください。(アクセス許可のプロンプトが表示された場合は、許可を実施してください)

3. 以下コマンドより、GASに紐づけれられているスプレッドシートを開くと、現時点の年月の予定がカウントされ稼動工数として記載されています。  

```bash
npx clasp open --addon
```

※もし対象の年月を指定したい場合、プロジェクトの「プロジェクト」より、スクリプトプロパティに以下の変数を設定してから`handler`関数を実行してください。

| 変数名 | 値 | 例 |
| --- | --- | --- |
| YEAR | 対象の年 | 2024 |
| MONTH | 対象の月 | 1 |
