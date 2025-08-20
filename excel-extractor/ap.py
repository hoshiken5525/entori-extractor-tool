import os
from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import openpyxl
import uuid

# Flaskアプリケーションの初期化
app = Flask(__name__)
# ファイルアップロード先の一時フォルダ
UPLOAD_FOLDER = 'temp_uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# アップロードフォルダが存在しない場合は作成
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- データ抽出ルールの定義 ---
# ここで指定されたルールに従ってExcelからデータを抽出します
EXTRACTION_RULES = [
    {'label': 'ディレクトリ名', 'cell': 'G30', 'type': 'cell'},
    {'label': '企業名', 'cell': 'C6', 'type': 'cell'},
    {'label': 'フリガナ', 'cell': 'C5', 'type': 'cell'},
    {'label': '代表者', 'cell': 'C8', 'type': 'cell', 'format': lambda s: s.replace('　', ' ') if s else s},
    {'label': '担当者', 'cell': 'C24', 'type': 'cell', 'format': lambda s: s.replace('　', ' ') if s else s},
    {'label': '郵便番号', 'cell': 'C10', 'type': 'cell', 'format': lambda s: s.replace('〒', '').strip() if s else s},
    {'label': '所在地詳細', 'cell': 'E12', 'type': 'cell', 'format': lambda s: s.replace('　', ' ') if s else s},
    {'label': 'TEL', 'cell': 'C7', 'type': 'cell', 'format': lambda s: s.replace('-', '').replace('ー', '') if s else s},
    {'label': 'FAX', 'cell': 'H7', 'type': 'cell', 'format': lambda s: s.replace('-', '').replace('ー', '') if s else s},
    {'label': '会社URL', 'cell': 'D13', 'type': 'cell', 'format': lambda s: f"https://{s}" if s and not s.startswith(('http://', 'https://')) else s},
    {'label': '企業メールアドレス', 'cell': 'G32', 'type': 'cell'},
    # --- 以下は固定テキスト ---
    {'label': '担当1', 'text': '清水隼人', 'type': 'fixed'},
    {'label': '担当2', 'text': '山賀徳能', 'type': 'fixed'},
    {'label': '担当3', 'text': '竹下章太郎', 'type': 'fixed'},
    {'label': '担当4', 'text': '成重裕樹', 'type': 'fixed'},
    {'label': '担当5', 'text': '石塚 恵', 'type': 'fixed'},
    {'label': '担当6', 'text': '千々和崇', 'type': 'fixed'},
]

# --- Webページの表示 (ルートURL) ---
@app.route('/')
def index():
    # index.htmlをブラウザに表示します
    return render_template('index.html')

# --- ファイルアップロードとExcel処理 ---
@app.route('/upload', methods=['POST'])
def upload_file():
    # ファイルがリクエストに含まれているか確認
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    file = request.files['file']
    # ファイル名が空でないか確認
    if file.filename == '':
        return jsonify({'error': 'ファイル名がありません'}), 400

    # 安全なファイル名を生成し、一時的に保存
    filename = secure_filename(str(uuid.uuid4()) + os.path.splitext(file.filename)[1])
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        file.save(filepath)

        # openpyxlでExcelファイルを開く
        workbook = openpyxl.load_workbook(filepath)
        # 常に最初のシートを対象とする
        sheet = workbook.active
        
        extracted_data = []
        # 定義されたルールに基づいてデータを一つずつ抽出
        for rule in EXTRACTION_RULES:
            value = ''
            # ルールが 'cell' の場合、セルから値を読み取る
            if rule['type'] == 'cell':
                cell_value = sheet[rule['cell']].value
                # セルが空でない場合のみ処理
                if cell_value is not None:
                    value = str(cell_value)
                    # フォーマット（データ加工）ルールがあれば適用
                    if 'format' in rule:
                        value = rule['format'](value)
            # ルールが 'fixed' の場合、定義された固定テキストを使用
            elif rule['type'] == 'fixed':
                value = rule['text']
            
            extracted_data.append({'label': rule['label'], 'value': value})

        # 抽出したデータをJSON形式で返す
        return jsonify(extracted_data)

    except Exception as e:
        # エラーが発生した場合
        return jsonify({'error': f"処理中にエラーが発生しました: {str(e)}"}), 500
    finally:
        # 処理が成功しても失敗しても、必ず一時ファイルを削除
        if os.path.exists(filepath):
            os.remove(filepath)

# --- アプリケーションの実行 ---
# この部分はローカルでのテスト実行にのみ使用されます。
# Renderなどの本番環境では、Gunicornというサーバーが直接 'app' 変数を起動します。
if __name__ == '__main__':
    # debug=True は開発時のみ使用し、公開時は必ず False にするか削除します
    app.run()
