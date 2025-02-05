from flask import Flask, jsonify, render_template, request
from flask_cors import CORS
import socket
import json
from dotenv import load_dotenv
from task_database import TaskDatabase
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
import os
import ollama
import traceback
import sys
import yaml
import base64
import imghdr
from PIL import Image
import re
from datetime import datetime
import pptx
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
from flask_restx import Api, Resource, fields

# Charger les variables d'environnement du fichier .env
load_dotenv()

# Configuration Flask
app = Flask(__name__)
CORS(app)

# Initialisation de la base de données
task_db = TaskDatabase('tasks.db')

# Configuration Ollama
ollama_config = {
    'host': os.getenv('OLLAMA_HOST', 'http://localhost:11434'),
    'model': os.getenv('OLLAMA_MODEL', 'llama3')
}

# Client Ollama
ollama_client = ollama.Client(ollama_config['host'])

# Définition de la fonction de conversion de couleur
def convert_color_to_rgb(color):
    color_map = {
        'rouge': [255, 0, 0],
        'bleu': [0, 0, 255],
        'vert': [0, 255, 0],
        'jaune': [255, 255, 0],
        'orange': [255, 165, 0],
        'violet': [128, 0, 128],
        'rose': [255, 192, 203],
        'marron': [165, 42, 42],
        'gris': [128, 128, 128],
        'noir': [0, 0, 0],
        'blanc': [255, 255, 255]
    }
    return color_map.get(color.lower(), [0, 0, 255])  # Bleu par défaut

# Définition de la fonction d'analyse du prompt
def parse_project_prompt(client, prompt, config):
    try:
        model = config.get('model', 'mervinpraison/llama3.2-3B-instruct-test-2:8b')
        
        system_prompt = """
        Tu es un assistant spécialisé dans l'analyse de prompts de projet.
        Tu dois identifier si le prompt est une création, une mise à jour ou une suppression de projet

        RÈGLES D'IDENTIFICATION DU TYPE D'ACTION :
        - "create" : Utiliser des mots comme "créer", "ajouter", "nouveau", "new", "initialiser"
        - "update" : Utiliser des mots comme "modifier", "changer", "update", "mettre à jour", "ajuster"
        - "delete" : Utiliser des mots comme "supprimer", "effacer", "delete", "remove", "retirer"

        EXEMPLES :
        1. Prompt "Créer projet 'P1' du 15 mai au 29 décembre" → type: "create"
        2. Prompt "Modifier le projet 'P1' pour changer ses dates" → type: "update"
        3. Prompt "Supprimer le projet 'P1'" → type: "delete"

        Extrait précisément les informations suivantes :

        1. NOM DU PROJET :
        - Extraire le nom exact entre guillemets ou apostrophes
        - Conserver la casse et les espaces originaux
        - Si absent, retourner "Unnamed Project"

        2. DATES :
        - Toujours convertir au format "YYYY/MM/DD"
        - Identifier les dates avec flexibilité : 
            * Formats acceptés : JJ/MM/AAAA, MM/JJ/AAAA, AAAA-MM-JJ
            * Mots-clés : "du", "from", "entre", "from...to"
        - Si une date manque, utiliser NULL

        3. COULEUR DU PROJET :
        - Accepter les formats :
            * Noms de couleurs (rouge, bleu, vert)
            * Valeurs RGB entre 0-255
            * Codes hexadécimaux
        - Conversion automatique en RGB
        - Si non spécifié, utiliser une couleur par défaut

        CONSEILS SUPPLÉMENTAIRES :
        - Soyez précis et littéral
        - En cas d'ambiguïté, choisissez l'interprétation la plus probable
        
        
        Réponds UNIQUEMENT au format JSON suivant avec tous les champs obligatoires :
        {
            "type": "create|update|delete",
            "task_name": "Nom du projet",
            "start_date": "YYYY/MM/DD",
            "end_date": "YYYY/MM/DD",
            "start_month": [index_mois, position_dans_mois],
            "end_month": [index_mois, position_dans_mois],
            "color_rgb": [R, G, B]
        }
        
        Règles pour le calcul de position_dans_mois :
        - 0.0 : début du mois (1-10)
        - 0.5 : milieu du mois (11-20)
        - 1.0 : fin du mois (21-31)

        Règles pour le calcul de index_mois :
        - Pour start_month : index_mois = start_date(mois) - 1
        - Pour end_month : index_mois = end_date(mois) - 1
        

        Exemples :
        Prompt: "je veux un projet 'P1' du 15 mai au 29 decembre (couleur : vert)"
        Réponse: {
            "type": "create",
            "task_name": "P1",
            "start_date": "2025/05/15",
            "end_date": "2025/12/29",
            "start_month": [4, 0.5],
            "end_month": [11, 1.0],
            "color_rgb": [0, 255, 0]
        }

        Prompt: "je veux créer une tâche 'P1' du 25 mars au 3 aout (couleur : vert)"
        Réponse: {
            "type": "create",
            "task_name": "P1",
            "start_date": "2025/03/25",
            "end_date": "2025/08/03",
            "start_month": [2, 1.0],
            "end_month": [7, 0.0],
            "color_rgb": [0, 255, 0]
        }        

        Prompt: "je veux modifier le projet 'P1' pour qu'il commence le 1er juin"
        Réponse: {
            "type": "update",
            "task_name": "P1",
            "start_date": "2025/06/01",
            "end_date": null,
            "start_month": [5, 0.0],
            "end_month": null,
            "color_rgb": null

        Prompt: "Change 'P1' pour pour avir une date de fin au 13 avril"
        Réponse: {
            "type": "update",
            "task_name": "P1",
            "start_date": null,
            "end_date": "2025/04/13",
            "start_month": null,
            "end_month": [3, 0.5],
            "color_rgb": null

        Prompt: "Update le projet 'P1' avec les dates début 25 mars et fin 3 aout (couleur : vert)"
        Réponse: {
            "type": "update",
            "task_name": "P1",
            "start_date": "2025/03/25",
            "end_date": "2025/08/03",
            "start_month": [2, 1.0],
            "end_month": [7, 0.0],
            "color_rgb": [0, 255, 0]

        Prompt: "Supprime le projet 'P1'"
        Réponse: {
            "type": "delete",
            "task_name": "P1",
            "start_date": null,
            "end_date": null,
            "start_month": null,
            "end_month": null,
            "color_rgb": null
        }
        """
        
        response = client.chat(
            model=model,
            messages=[
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': prompt}
            ]
        )
        
        result = response['message']['content'].strip()

        try:
            start_index = result.find('{')
            end_index = result.rfind('}') + 1
            
            if start_index != -1 and end_index != -1:
                result = result[start_index:end_index]
            
            parsed_res = json.loads(result)
            
            required_keys = ['type', 'task_name']
            if not all(key in parsed_res for key in required_keys):
                raise ValueError("JSON incomplet")
            
            if parsed_res['type'] == 'update':
                parsed_res = {k: v for k, v in parsed_res.items() if v is not None}
                                        
            return parsed_res
        
        except json.JSONDecodeError as e:
            print(f"Erreur de décodage JSON : {e}")
            print(f"Contenu problématique : {result}")
            return None
        except ValueError as e:
            print(f"Erreur de validation JSON : {e}")
            return None
    
    except Exception as e:
        print(f"Erreur lors de l'analyse du prompt : {e}")
        return None

# Définition de la fonction de création de tâche sur la roadmap
def create_task_on_roadmap(prs, task_info):
    task_name = task_info['task_name']
    
    start_month = task_info['start_month'][0] if isinstance(task_info['start_month'], (list, tuple)) else task_info['start_month']
    start_pos = task_info['start_month'][1] if isinstance(task_info['start_month'], (list, tuple)) else task_info['start_month']
    
    end_month = task_info['end_month'][0] if isinstance(task_info['end_month'], (list, tuple)) else task_info['end_month']
    end_pos = task_info['end_month'][1] if isinstance(task_info['end_month'], (list, tuple)) else task_info['end_month']
    
    start_month = 0 if start_month is None else start_month
    end_month = 11 if end_month is None else end_month
    
    start_pos = 0.5 if start_pos is None else start_pos
    end_pos = 1.0 if end_pos is None else end_pos
    
    color_rgb = task_info['color_rgb']         # Couleur RGB

    roadmap_slide = prs.slides[0]  # Première slide (roadmap)

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    grid_margin_top = Inches(1.5)  # Marge en haut
    grid_margin_bottom = Inches(0.5)  # Marge en bas
    grid_margin_left = Inches(0.5)  # Marge à gauche
    grid_margin_right = Inches(0.5)  # Marge à droite

    grid_height = slide_height - grid_margin_top - grid_margin_bottom
    grid_width = slide_width - grid_margin_left - grid_margin_right

    month_width = grid_width / 12

    start_x = grid_margin_left + (start_month * month_width) + (start_pos * month_width)
    end_x = grid_margin_left + (end_month * month_width) + (end_pos * month_width)

    task_height = Inches(0.5)
    
    existing_shapes = [shape for shape in roadmap_slide.shapes if shape.has_text_frame]
    
    task_y = grid_margin_top + Inches(0.5) + (len(existing_shapes) * Inches(0.6))

    task_shape = roadmap_slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, 
        start_x, 
        task_y, 
        end_x - start_x, 
        task_height
    )

    task_shape.fill.solid()
    task_shape.fill.fore_color.rgb = RGBColor(color_rgb[0], color_rgb[1], color_rgb[2])
    task_shape.line.fill.background()

    text_frame = task_shape.text_frame
    text_frame.text = task_name
    text_frame.paragraphs[0].font.size = Pt(10)
    text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)  # Texte en noir
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Définition de la fonction de création de slide de roadmap
def create_roadmap_slide(prs, task_info=None):
    if len(prs.slides) > 0:
        slide = prs.slides[0]
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Layout vide
    
    months_grid_exists = any(shape.has_table for shape in slide.shapes)
    
    if not any(shape.has_text_frame and shape.text_frame.text == "ROADMAP" for shape in slide.shapes):
        title = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.5))
        title.text_frame.text = "ROADMAP"
        title.text_frame.paragraphs[0].font.size = Pt(24)
        title.text_frame.paragraphs[0].font.bold = True
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)    
    
    if not months_grid_exists:
        print("Ajout de la grille des mois")
        
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        margin_left = Inches(0.5)
        margin_right = Inches(0.5)
        
        effective_width = slide_width - (margin_left + margin_right)
        
        months_box = slide.shapes.add_table(
            2,  # 2 rangées
            12,  # 12 colonnes (mois)
            margin_left,  # Position X de départ
            Inches(1.5),  # Position Y
            effective_width,  # Largeur totale
            Inches(0.5)  # Hauteur
        ).table
        
        months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        
        for i, month in enumerate(months):
            cell = months_box.cell(0, i)
            cell.text = month
            cell.text_frame.paragraphs[0].font.size = Pt(8)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    else:
        print("Grille des mois déjà existante")
    
    if task_info:
        create_task_on_roadmap(prs, task_info)
    
    return slide

# Définition de la fonction de conversion de tâche de la base de données en task_info
def convert_db_task_to_task_info(db_task):
    color_rgb = json.loads(db_task['color_rgb']) if db_task['color_rgb'] else None
    
    task_info = {
        "type": "create",  # Par défaut, toujours "create"
        "task_name": db_task['task_name'],
        "start_month": [
            db_task['start_month'] if db_task['start_month'] is not None else None,
            db_task['start_position'] if db_task['start_position'] is not None else 0.5
        ],
        "end_month": [
            db_task['end_month'] if db_task['end_month'] is not None else None,
            db_task['end_position'] if db_task['end_position'] is not None else 1.0
        ],
        "color_rgb": color_rgb
    }
    
    return task_info

# Définition de la fonction de normalisation de texte
def normalize_text(text):
    text = text.lower()
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)  # Remplacer les espaces multiples par un seul
    text = re.sub(r'[^\w\s]', '', text)  # Supprimer la ponctuation
    
    return text

# Définition de la fonction de liste des objets PowerPoint
def list_powerpoint_objects(file_path):
    if not os.path.exists(file_path):
        print(f"Erreur : Le fichier {file_path} n'existe pas.")
        return []

    prs = Presentation(file_path)
    
    ignored_objects = ['ROADMAP']
    
    objects_list = []
    
    for slide_index, slide in enumerate(prs.slides, 1):
        slide_objects = []
        
        for shape_index, shape in enumerate(slide.shapes, 1):
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                
                if text and text not in ignored_objects:
                    slide_objects.append({
                        'type': 'texte',
                        'slide': slide_index,
                        'index': shape_index,
                        'text': text
                    })
            
            elif shape.has_table:
                table = shape.table
                first_row_texts = [cell.text.strip() for cell in table.rows[0].cells]
                if not (len(first_row_texts) > 1 and all(len(month) == 3 for month in first_row_texts)):
                    table_rows = []
                    for row_index, row in enumerate(table.rows, 1):
                        row_texts = [cell.text.strip() for cell in row.cells]
                        if any(row_texts):
                            table_rows.append({
                                'row_index': row_index,
                                'row_texts': row_texts
                            })
                    
                    slide_objects.append({
                        'type': 'tableau',
                        'slide': slide_index,
                        'index': shape_index,
                        'rows': table_rows
                    })
        
        if slide_objects:
            objects_list.extend(slide_objects)
    
    return [normalize_text(obj['text']) for obj in objects_list if obj['type'] == 'texte']

# Définition de la fonction de traitement de ligne de prompt
def process_prompt_line(prompt_line):
    templates_dir = "templates"
    output_dir = "generated"
    
    os.makedirs(templates_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    template_path = os.path.join(templates_dir, "roadmap.pptx")
    output_path = os.path.join(output_dir, "roadmap.pptx")
    
    config = {
        'host': os.getenv('OLLAMA_HOST', 'http://localhost:11434'),
        'model': os.getenv('OLLAMA_MODEL', 'llama3')
    }
    client = ollama.Client(config['host'])
    
    print(f"\n--- Traitement du prompt : {prompt_line} ---")
    
    try:
        task_info = parse_project_prompt(client, prompt_line, config)
        
        if task_info and task_info.get('type') in ['create', 'update']:
            task_id = task_db.upsert_task(task_info, raw_prompt=prompt_line)
            print(f"Tâche créée ou mise à jour avec l'ID : {task_info}")
        elif task_info and task_info.get('type') == 'delete':
            task_db.delete_task(task_info, raw_prompt=prompt_line)
            print(f"Tâche supprimée : {task_info.get('task_name')}")
        
        if os.path.exists(template_path):
            prs = Presentation(template_path)
        else:
            prs = Presentation()
            prs.save(template_path)
        
        all_tasks = task_db.list_tasks()
        
        while len(prs.slides) > 0:
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
        
        for task in all_tasks:
            task_info = convert_db_task_to_task_info(task)
            
            create_roadmap_slide(prs, task_info)
        
        prs.save(output_path)
        print(f"Présentation mise à jour : {output_path}")
        
        return task_info
    
    except Exception as e:
        print(f"Erreur lors du traitement du prompt '{prompt_line}' : {e}")
        traceback.print_exc()
        return None

# Définition de la fonction de traitement de prompt
def process_prompt(body):
    try:
        prompt = body.get('prompt')
        if not prompt:
            return {'error': 'Prompt manquant'}, 400
        
        task_info = process_prompt_line(prompt)
        
        update_presentation()
        
        return task_info, 200
    except Exception as e:
        return {'error': str(e)}, 500

# Définition de la fonction de mise à jour de la présentation
def update_presentation():
    templates_dir = "templates"
    output_dir = "generated"
    
    template_path = os.path.join(templates_dir, "roadmap.pptx")
    output_path = os.path.join(output_dir, "roadmap.pptx")
    
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()
        prs.save(template_path)
    
    all_tasks = task_db.list_tasks()
    
    while len(prs.slides) > 0:
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
    
    for task in all_tasks:
        task_info = convert_db_task_to_task_info(task)
        
        create_roadmap_slide(prs, task_info)
    
    prs.save(output_path)
    print(f"Présentation mise à jour : {output_path}")

# Fonction pour trouver un port libre
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

# Configuration de l'API
api = Api(app,
          version='1.0',
          title='Roadmap API',
          doc='/swagger-ui/',
          validate=True)

# Modèle de données
prompt_model = api.model('Prompt', {
    'prompt': fields.String(required=True, description='Description textuelle du projet')
})

# Namespace API
ns = api.namespace('projects', description='Opérations sur les projets')

@ns.route('/process_prompt')
class ProcessPrompt(Resource):
    @ns.expect(prompt_model)
    @ns.response(200, 'Success')
    @ns.response(400, 'Invalid request')
    def post(self):
        body = api.payload
        
        if 'prompt' not in body:
            api.abort(400, 'Prompt manquant')
        
        prompt = body['prompt']
        
        try:
            # ... logique existante de process_prompt() ...
            task_info = process_prompt_line(prompt)
            return {'message': 'Projet créé', 'task': task_info}, 200
        except Exception as e:
            return {'error': str(e)}, 500

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/tasks')
def get_tasks():
    tasks = task_db.list_tasks()
    lanes = {}
    processed_tasks = []
    
    for task in tasks:
        start = task['start_month']
        end = task['end_month']
        lane = 0
        
        while any(t['end_month'] > start and lanes.get(t['id'], -1) == lane 
                 for t in processed_tasks):
            lane += 1
        
        processed_tasks.append({**task, 'lane': lane})
        lanes[task['id']] = lane
    
    return jsonify([{
        'task_name': t['task_name'],
        'start_percent': (t['start_month']/12)*100,
        'duration_percent': ((t['end_month']-t['start_month'])/12)*100,
        'color_rgb': json.loads(t['color_rgb']),
        'lane': t['lane']
    } for t in processed_tasks])

# Lancement de l'application
if __name__ == '__main__':
    free_port = find_free_port()
    print(f"Démarrage du serveur sur le port {free_port}")
    app.run(port=free_port, debug=True)