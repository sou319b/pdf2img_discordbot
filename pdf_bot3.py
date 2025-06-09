import discord
import os
from pdf2image import convert_from_path
from dotenv import load_dotenv
import subprocess # libreoffice呼び出しのため追加
import tempfile # 一時ディレクトリ管理のため追加

load_dotenv() # .envファイルから環境変数を読み込む

# Discordボットのアクセストークンを環境変数から取得
DISCORD_BOT_TOKEN = os.getenv('DISCORD_BOT_TOKEN')

# Popplerのパスを環境変数から取得 (Windowsユーザー向け)
POPPLER_PATH = os.getenv('POPPLER_PATH')

# 対応するOfficeドキュメントの拡張子
OFFICE_EXTENSIONS = ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx')

intents = discord.Intents.default()
intents.message_content = True # メッセージ内容へのアクセスを有効にする

client = discord.Client(intents=intents)

@client.event
async def on_ready():
    print(f'{client.user}としてログインしました')

async def convert_office_to_pdf(office_path, output_dir):
    """
    LibreOfficeを使用してOfficeドキュメントをPDFに変換する。
    変換されたPDFのパスを返す。失敗した場合はNoneを返す。
    """
    try:
        # libreofficeコマンドが --headless モードでPDFに変換し、指定されたディレクトリに出力
        #libreoffice --headless --convert-to pdf input_file.docx --outdir /path/to/output
        #libreoffice --infilter="Microsoft Word 2007/2010/2013 XML" --headless --convert-to pdf ./test_data/test.docx  --outdir ./test_data
        current_env = os.environ.copy()
        current_env["LANG"] = "ja_JP.UTF-8"
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', office_path, '--outdir', output_dir],
            capture_output=True, text=True, check=True, timeout=60, env=current_env # envパラメータを追加
        )
        # 変換後のPDFファイル名は元ファイル名.pdfとなる
        pdf_filename = os.path.splitext(os.path.basename(office_path))[0] + '.pdf'
        converted_pdf_path = os.path.join(output_dir, pdf_filename)
        if os.path.exists(converted_pdf_path):
            return converted_pdf_path
        else:
            print(f"PDF変換エラー: {converted_pdf_path} が見つかりません。libreofficeの出力: {result.stdout} {result.stderr}")
            return None
    except subprocess.CalledProcessError as e:
        print(f"libreofficeコマンドの実行に失敗しました: {e}")
        print(f"Stderr: {e.stderr}")
        print(f"Stdout: {e.stdout}")
        return None
    except subprocess.TimeoutExpired:
        print(f"libreofficeのPDF変換がタイムアウトしました: {office_path}")
        return None
    except Exception as e:
        print(f"OfficeからPDFへの変換中に予期せぬエラーが発生しました: {e}")
        return None

@client.event
async def on_message(message):
    # ボット自身のメッセージは無視
    if message.author == client.user:
        return

    # 処理対象の添付ファイルを選別 (PDFまたは指定Office形式)
    attachments_to_process = []
    for att in message.attachments:
        filename_lower = att.filename.lower()
        if filename_lower.endswith('.pdf') or filename_lower.endswith(OFFICE_EXTENSIONS):
            attachments_to_process.append(att)

    if not attachments_to_process:
        # '!hello' のような他のコマンドはこちらで処理
        if message.content.startswith('!hello'):
            await message.channel.send('こんにちは！')
        return


    # スレッド名を生成
    thread_name = f"{message.author.display_name}さんのファイル画像変換"
    thread = None
    try:
        thread = await message.create_thread(name=thread_name)
        await thread.send(f"{len(attachments_to_process)}個のファイルを画像に変換します...")
    except discord.HTTPException as e:
        print(f"スレッドの作成に失敗しました: {e}")
        await message.channel.send(f"{len(attachments_to_process)}個のファイルを画像に変換します... (スレッド作成に失敗したため、このチャンネルに投稿します)")
    except discord.Forbidden:
        print("スレッドを作成する権限がありません。")
        await message.channel.send(f"{len(attachments_to_process)}個のファイルを画像に変換します... (スレッド作成権限がないため、このチャンネルに投稿します)")

    target_channel = thread if thread else message.channel

    # 一時ファイルを保存するディレクトリを作成
    with tempfile.TemporaryDirectory() as temp_dir:
        for attachment in attachments_to_process:
            original_file_path = os.path.join(temp_dir, attachment.filename)
            pdf_to_convert_path = None # 実際に画像変換するPDFのパス
            temp_office_pdf_path = None # Officeから変換された一時PDFのパス

            try:
                await attachment.save(original_file_path)
                await target_channel.send(f'{attachment.filename} を受信しました。処理中です...')

                filename_lower = attachment.filename.lower()
                if filename_lower.endswith(OFFICE_EXTENSIONS):
                    # OfficeファイルをPDFに変換
                    converted_pdf = await convert_office_to_pdf(original_file_path, temp_dir)
                    if converted_pdf and os.path.exists(converted_pdf):
                        pdf_to_convert_path = converted_pdf
                        temp_office_pdf_path = converted_pdf # 後で削除するために保持
                        await target_channel.send(f'{attachment.filename} をPDFに変換しました: {os.path.basename(converted_pdf)}')
                    else:
                        await target_channel.send(f'{attachment.filename} のPDFへの変換に失敗しました。')
                        continue # 次のファイルへ
                elif filename_lower.endswith('.pdf'):
                    pdf_to_convert_path = original_file_path
                else: # ここには来ないはずだが念のため
                    await target_channel.send(f'{attachment.filename} は対応していないファイル形式です。')
                    continue


                if not pdf_to_convert_path:
                    continue

                # PDFを画像に変換
                image_paths = []
                if POPPLER_PATH: # Windows環境向けPopplerパス指定
                    images = convert_from_path(pdf_to_convert_path, poppler_path=POPPLER_PATH)
                else: # Linux/macOSなどPopplerがPATHにあればこちら
                    images = convert_from_path(pdf_to_convert_path)

                if not images:
                    await target_channel.send(f'{attachment.filename} (PDF: {os.path.basename(pdf_to_convert_path)}) から画像を抽出できませんでした。')
                    continue

                for i, image in enumerate(images):
                    base_filename = os.path.splitext(attachment.filename)[0]
                    image_path = os.path.join(temp_dir, f'{base_filename}_page_{i}.png')
                    image.save(image_path, 'PNG')
                    image_paths.append(image_path)

                if image_paths:
                    chunk_size = 10
                    for i in range(0, len(image_paths), chunk_size):
                        chunk = image_paths[i:i + chunk_size]
                        files_to_send = [discord.File(fp=img_path) for img_path in chunk]
                        if files_to_send:
                            if i == 0:
                                await target_channel.send(f'{attachment.filename} の変換画像：', files=files_to_send)
                            else:
                                await target_channel.send(files=files_to_send)
                else:
                    await target_channel.send(f'{attachment.filename} の画像の変換または保存に失敗しました。')

            except Exception as e:
                print(f'{attachment.filename} の処理中にエラーが発生しました: {e}')
                await target_channel.send(f'{attachment.filename} の処理中にエラーが発生しました: {e}')
            finally:
                # original_file_path は temp_dir が閉じる際に自動で消える
                # temp_office_pdf_path (Officeから変換したPDF) と image_paths (生成画像) は
                # temp_dir 内に作られるので、これらも自動で消える。
                # 個別に os.remove する必要はない。
                pass # temp_dirが自動でクリーンアップ

    # PDF以外の添付ファイルのみの場合、または添付ファイルがない場合のコマンド処理は上に移動

if DISCORD_BOT_TOKEN:
    client.run(DISCORD_BOT_TOKEN)
else:
    print("エラー: DISCORD_BOT_TOKENが設定されていません。 .envファイルを確認してください。")

