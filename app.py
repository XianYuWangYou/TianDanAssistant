from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import json
from core import DocumentProcessor
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='.', static_url_path='')

# 配置上传文件夹
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'docs'
ALLOWED_EXTENSIONS = {'docx', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# 确保上传和输出文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/placeholders', methods=['POST'])
def extract_placeholders():
    try:
        # 获取上传的文件
        files = request.files.getlist('files')
        template_files = []
        
        # 保存上传的文件
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                template_files.append(file_path)
        
        # 使用DocumentProcessor提取占位符
        processor = DocumentProcessor()
        placeholders, placeholder_files = processor.collect_all_placeholders(template_files)
        
        # 清理上传的文件
        for file_path in template_files:
            if os.path.exists(file_path):
                os.remove(file_path)
        
        return jsonify({
            'success': True,
            'placeholders': list(placeholders),
            'placeholder_files': placeholder_files
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/process', methods=['POST'])
def process_documents():
    try:
        # 获取JSON数据
        data = request.get_json()
        template_files = data.get('template_files', [])
        user_inputs = data.get('user_inputs', {})
        output_dir = app.config['OUTPUT_FOLDER']
        
        # 复制模板文件到上传文件夹
        saved_files = []
        for filename in template_files:
            # 在实际应用中，这里需要处理文件上传
            # 现在我们假设文件已经存在于uploads文件夹中
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if os.path.exists(file_path):
                saved_files.append(file_path)
        
        # 处理文档
        processor = DocumentProcessor()
        generated_files = processor.process_templates(saved_files, user_inputs, output_dir)
        
        return jsonify({
            'success': True,
            'generated_files': generated_files
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/schemes', methods=['GET', 'POST'])
def schemes():
    data_file = 'app_data.json'
    
    if request.method == 'GET':
        # 读取方案数据
        if os.path.exists(data_file):
            with open(data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                schemes = data.get('schemes', {})
                return jsonify({
                    'success': True,
                    'schemes': schemes
                })
        else:
            return jsonify({
                'success': True,
                'schemes': {}
            })
    
    elif request.method == 'POST':
        # 保存方案数据
        try:
            scheme_data = request.get_json()
            schemes_data = {}
            
            # 读取现有数据
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    schemes_data = json.load(f)
            
            # 确保schemes键存在
            if 'schemes' not in schemes_data:
                schemes_data['schemes'] = {}
            
            # 更新方案数据
            schemes_data['schemes'].update(scheme_data)
            
            # 保存数据
            with open(data_file, 'w', encoding='utf-8') as f:
                json.dump(schemes_data, f, ensure_ascii=False, indent=2)
            
            return jsonify({
                'success': True
            })
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)