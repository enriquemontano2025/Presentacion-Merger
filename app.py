from flask import Flask, render_template, request, jsonify, send_file
import os
import json
from pathlib import Path
from utils.presentation_merger import PresentationMerger
import tempfile
import shutil

app = Flask(__name__)
app.config['SECRET_KEY'] = 'jll-presentation-merger-2024'

# Configuración
UPLOAD_FOLDER = 'static/uploads'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scan-folder', methods=['POST'])
def scan_folder():
    """Escanea una carpeta y devuelve las presentaciones disponibles"""
    data = request.get_json()
    folder_path = data.get('folder_path', '')
    
    if not folder_path or not os.path.exists(folder_path):
        return jsonify({'error': 'Carpeta no válida o no existe'}), 400
    
    try:
        merger = PresentationMerger(folder_path)
        presentations = merger.scan_presentations()
        return jsonify({
            'success': True,
            'presentations': presentations,
            'total': len(presentations['ESP']) + len(presentations['ENG'])
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/merge', methods=['POST'])
def merge_presentations():
    try:
        data = request.get_json()
        folder_path = data.get('folder_path')
        language = data.get('language', 'ESP')
        buildings = data.get('buildings', [])
        output_name = data.get('output_name', 'Propuesta_JLL')
        
        if not folder_path or not buildings:
            return jsonify({'error': 'Faltan parámetros requeridos'}), 400
        
        # Crear carpeta de uploads si no existe
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        
        # Crear instancia del merger
        merger = PresentationMerger(folder_path, language)
        
        # Generar archivo
        output_path = os.path.join(UPLOAD_FOLDER, f"{output_name}_{language}.pptx")
        
        # Realizar la combinación
        success, message = merger.merge_presentations(buildings, output_path)
        
        if success:
            return jsonify({
                'success': True,
                'message': message,
                'download_url': f'/download/{os.path.basename(output_path)}',
                'filename': os.path.basename(output_path)
            })
        else:
            return jsonify({'error': message}), 400
            
    except Exception as e:
        return jsonify({'error': f'Error: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Descarga el archivo generado"""
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return "Archivo no encontrado", 404
    except Exception as e:
        return f"Error al descargar: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
