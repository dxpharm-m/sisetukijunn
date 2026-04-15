"""
調剤報酬 施設基準 届出書類作成ツール - バックエンドAPI
"""
import os, json, zipfile, tempfile, shutil, traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS

BASE = Path(__file__).parent
app = Flask(__name__)
CORS(app)

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/')
def index():
    index_file = BASE / 'index.html'
    if not index_file.exists():
        return f'index.html not found in {BASE}. Files: {list(BASE.glob("*.html"))}', 404
    content = index_file.read_text(encoding='utf-8')
    return Response(content, mimetype='text/html; charset=utf-8')

@app.route('/api/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'データが空です'}), 400

        out_dir = Path(tempfile.mkdtemp())
        try:
            import sys
            sys.path.insert(0, str(BASE))
            from fill_forms import generate_all
            results = generate_all(data, str(out_dir))

            generated = [r for r in results if r['ok']]
            if not generated:
                return jsonify({'error': '書類の生成に失敗しました'}), 500

            basic    = data.get('basic', {})
            nm       = basic.get('b_name', '薬局').replace(' ','').replace('　','')
            dt       = basic.get('b_date', '').replace('-','')
            zip_name = f'施設基準届出書類_{nm}_{dt}.zip'
            zip_path = out_dir / zip_name

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for r in generated:
                    fp = out_dir / r['file']
                    if fp.exists():
                        zf.write(fp, r['file'])

            return send_file(str(zip_path), as_attachment=True,
                           download_name=zip_name, mimetype='application/zip')
        finally:
            pass

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'サーバーエラー: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
