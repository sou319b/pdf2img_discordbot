# PDF/Office ファイル画像変換 Discordボット

Discordに添付されたPDFファイルやOfficeドキュメント（Word, Excel, PowerPoint）を自動で画像（PNG形式）に変換するBotです。

ファイルが添付されたメッセージを検知すると、Botが自動でスレッドを作成し、その中でページごとに画像化したファイルを投稿します。

## 主な機能

* **PDFファイルの画像化**: PDFの各ページをPNG画像に変換します。
* **Officeファイルの画像化**: Word(.doc, .docx), Excel(.xls, .xlsx), PowerPoint(.ppt, .pptx) ファイルを一度PDFに変換し、その後画像化します。
* **自動スレッド作成**: 変換処理はファイル投稿者ごとのスレッド内で実行されるため、チャンネルが散らかりません。
* **複数ファイル対応**: 一つのメッセージに添付された複数のファイルも同時に処理します。

## 動作要件

このBotを動作させるには、以下のソフトウェアがインストールされている必要があります。

* Python 3.8 以上
* **LibreOffice**: OfficeドキュメントをPDFに変換するために使用します。
* **Poppler**: PDFを画像に変換するために使用します。

## セットアップ手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/sou319b/pdf2img_discordbot.git
cd pdf2img_discordbot
```

### 2. 外部ソフトウェアのインストール

お使いのOSに合わせて、LibreOfficeとPopplerをインストールしてください。

**Windows の場合:**

1.  **LibreOffice**: [公式サイト](https://ja.libreoffice.org/download/download/)からダウンロードしてインストールします。
2.  **Poppler**: [このページ](https://github.com/oschwartz10612/poppler-windows/releases/)などからダウンロードし、解凍して任意のフォルダ（例: `C:\poppler`）に配置します。その後、`C:\poppler\bin` のように、binフォルダへのパスをシステムの環境変数 `Path` に追加してください。
    * 環境変数 `Path` に追加しない場合は、後述の `.env` ファイルで `POPPLER_PATH` を指定する必要があります。

**macOS の場合 (Homebrew使用):**

```bash
# LibreOffice (Cask)
brew install --cask libreoffice

# Poppler
brew install poppler
```

**Debian/Ubuntu の場合:**

```bash
sudo apt update
sudo apt install -y libreoffice poppler-utils
```

### 3. Python環境のセットアップ

```bash
# 仮想環境の作成 (推奨)
python -m venv venv

# 仮想環境のアクティベート
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate

# 必要なライブラリのインストール
pip install -r requirements.txt
```

### 4. Discord Bot の設定

1.  [Discord Developer Portal](https://discord.com/developers/applications) にアクセスし、新しいアプリケーションを作成します。
2.  作成したアプリケーションの「Bot」タブから「Add Bot」をクリックしてBotを作成します。
3.  `TOKEN` をコピーしておきます。
4.  以下の「Privileged Gateway Intents」を **有効** にしてください。
    * `MESSAGE CONTENT INTENT`
5.  「OAuth2 > URL Generator」タブで、以下のスコープと権限を設定して招待URLを生成し、Botをサーバーに招待します。
    * **Scopes**:
        * `bot`
    * **Bot Permissions**:
        * `Send Messages`
        * `Read Message History`
        * `Create Public Threads` (または `Create Private Threads`)
        * `Send Messages in Threads`
        * `Attach Files`

### 5. 環境変数ファイル (.env) の作成

プロジェクトのルートディレクトリに `.env` という名前のファイルを作成し、以下のように設定を記述します。

```env
# Discord Botのアクセストークン
DISCORD_BOT_TOKEN=ここにBotのトークンを貼り付け

# 【Windowsユーザー向け】Popplerのパス
# Path環境変数にPopplerのbinディレクトリを追加していない場合、以下のようにフルパスを指定してください。
# 例: POPPLER_PATH=C:\path\to\poppler-23.11.0\Library\bin
POPPLER_PATH=
```

### 6. Botの起動

```bash
python pdf_bot3.py
```

コンソールに `(Bot名)としてログインしました` と表示されれば成功です。

## 使い方

1.  Botを導入したDiscordサーバーのチャンネルに、PDFまたはOfficeファイル（.docx, .xlsx, .pptxなど）を添付してメッセージを投稿します。
2.  Botが自動でスレッドを作成し、変換処理を開始します。
3.  処理が完了すると、スレッド内に画像ファイルが投稿されます。

## 注意事項

* 変換するファイルのサイズやページ数によっては、処理に時間がかかる場合があります。
* LibreOfficeのプロセスがタイムアウトした場合（デフォルト60秒）、処理は失敗します。
* 暗号化されたPDFや破損しているファイルは変換できません。

