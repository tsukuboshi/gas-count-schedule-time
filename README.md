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

- Node.js：v20以降
- TypeScript：v5以降

また、GASを実行するためのGoogleアカウントが必要です。  

## デプロイ手順

1. リポジトリをクローンし、ディレクトリを移動

```bash
git clone https://github.com/tsukuboshi/gas-count-schedule-time.git
cd gas-count-schedule-time
```

2. npmパッケージをインストール

```bash
npm init -y
npm install
```

3. asideを初期化しGASプロジェクトを作成

- Project Title：(任意のプロジェクト名)
- `license-header.txt`の上書き：No
- `rollup.config.mjs`の上書き：No
- Script ID：空文字
- Script ID for production environment：空文字

```
npx @google/aside init
✔ Project Title: … gas-count-schedule-time
✔ Adding scripts...
✔ Saving package.json...
✔ Installing dependencies...
license-header.txt already exists
✔ Overwrite … No
rollup.config.mjs already exists
✔ Overwrite … No
✔ Script ID (optional): … 
✔ Script ID for production environment (optional): … 
✔ Creating gas-count-schedule-time...

-> Google Sheets Link: https://drive.google.com/open?id=xxx
-> Apps Script Link: https://script.google.com/d/xxx/edit
```

4. GASアプリをデプロイ

```bash
npm run deploy
```
