import os
import sys

# Ajouter le dossier parent au chemin Python
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

import argparse
from core.template_processor import TemplateProcessor
from core.llm_integration import process_prompt
import yaml
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

# Définition des chemins de base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

app = Flask(__name__, static_folder='static')
CORS(app)  # Activer CORS pour tous les endpoints

def load_config():
    with open('config.yaml') as f:
        return yaml.safe_load(f)

@app.route('/')
def index():
    # Afficher les chemins de recherche
    print("BASE_DIR:", BASE_DIR)
    print("TEMPLATE_DIR:", TEMPLATE_DIR)
    print("Contenu de TEMPLATE_DIR:", os.listdir(TEMPLATE_DIR))
    
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Générateur de Roadmap</title>
    </head>
    <body>
        <h1>🛠 Générateur de Roadmap</h1>
        <form>
            <textarea placeholder="Entrez votre prompt..."></textarea>
            <button>Générer</button>
        </form>
    </body>
    </html>
    """
    return html_content

@app.route('/process_prompt', methods=['POST'])
def handle_process_prompt():
    data = request.json
    prompt = data.get('prompt')
    
    if not prompt:
        return jsonify({"message": "Prompt manquant"}), 400
    
    try:
        config = load_config()
        updates = process_prompt(prompt)
        
        processor = TemplateProcessor(config['template_path'])
        
        for slide_update in updates['slides']:
            processor.update_slide(slide_update['id'], slide_update['updates'])
        
        output_filename = f"roadmap_{hash(prompt)}.pptx"
        output_path = os.path.join(config['output_dir'], output_filename)
        processor.save(output_path)
        
        return jsonify({
            "message": "Présentation générée avec succès",
            "task": {
                "prompt": prompt,
                "output_file": output_filename
            }
        }), 200
    
    except Exception as e:
        return jsonify({"message": str(e)}), 500

def cli_main():
    parser = argparse.ArgumentParser(description='Automate PowerPoint updates')
    parser.add_argument('prompt', help='User modification prompt')
    parser.add_argument('-o', '--output', required=True, help='Output file name')
    args = parser.parse_args()
    
    config = load_config()
    updates = process_prompt(args.prompt)
    processor = TemplateProcessor(config['template_path'])
    
    for slide_update in updates['slides']:
        processor.update_slide(slide_update['id'], slide_update['updates'])
    
    output_path = os.path.join(config['output_dir'], args.output)
    processor.save(output_path)
    print(f"Présentation générée : {output_path}")

def run_server(port=5000):
    app.run(debug=True, port=port)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Automate PowerPoint updates or run server')
    parser.add_argument('-s', '--server', action='store_true', help='Run server')
    parser.add_argument('-p', '--port', type=int, default=5000, help='Server port')
    parser.add_argument('prompt', nargs='?', help='User modification prompt')
    parser.add_argument('-o', '--output', nargs='?', required=False, help='Output file name')
    args = parser.parse_args()
    
    if args.server:
        run_server(args.port)
    else:
        cli_main()
