import os
from dotenv import load_dotenv
from ollama import Client
import yaml
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import base64
import imghdr
from PIL import Image
import re
from datetime import datetime
import json
import traceback
import sys

# Charger les variables d'environnement du fichier .env
load_dotenv()

def load_config():
    with open('config.yaml', 'r') as f:
        return yaml.safe_load(f)

def encode_image_to_base64(image_path):
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"L'image {image_path} n'existe pas. Veuillez vérifier le chemin.")
    
    # Vérifier la taille et le type de fichier
    file_size = os.path.getsize(image_path)
    file_type = imghdr.what(image_path)
    
    print(f"Fichier image : {image_path}")
    print(f"Taille du fichier : {file_size} octets")
    print(f"Type d'image : {file_type}")
    
    if file_size > 20 * 1024 * 1024:  # Limite à 20 Mo
        raise ValueError("La taille de l'image est trop grande (> 20 Mo)")
    
    if file_type not in ['jpeg', 'png', 'gif', 'bmp']:
        raise ValueError(f"Type de fichier non supporté : {file_type}")
    
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def analyze_image_with_vision(client, image_path):
    try:
        base64_image = encode_image_to_base64(image_path)
        
        response = client.chat(
            model="llama3.2",  
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Analyse cette roadmap et donne-moi les informations suivantes en format JSON :\n"
                                   "1. Les mois présents\n"
                                   "2. Les tâches avec leurs périodes\n"
                                   "3. Les couleurs utilisées (en RGB)\n"
                                   "4. La légende (Done, Coming, Not planned)\n"
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=1000
        )
        return response.get('message', {}).get('content', '')
    except Exception as e:
        print(f"Erreur détaillée lors de l'analyse de l'image : {type(e).__name__} - {str(e)}")
        return None

def parse_date_to_month_position(client, date_str):
    """
    Convertit une date en français en position de mois.
    Gère différents formats de dates.
    """
    # Normalisation de la date
    date_str = date_str.lower().strip()
    
    # Dictionnaire de mapping des mois
    mois_mapping = {
        'janvier': 1, 'jan': 1, 
        'février': 2, 'fevrier': 2, 'fev': 2, 
        'mars': 3, 
        'avril': 4, 'avr': 4, 
        'mai': 5, 
        'juin': 6, 
        'juillet': 7, 'juil': 7, 
        'août': 8, 'aout': 8, 
        'septembre': 9, 'sept': 9, 
        'octobre': 10, 'oct': 10, 
        'novembre': 11, 'nov': 11, 
        'décembre': 12, 'decembre': 12, 'dec': 12
    }
    
    # Extraction du jour et du mois
    try:
        # Gestion des formats : 
        # 1. "1 mars", "1er mars", "01 mars"
        # 2. "2023-04-20", "01-01-2024"
        # 3. "20 avril 2023"
        
        # Formats avec année
        date_format_patterns = [
            r'(\d{4})-(\d{2})-(\d{2})',  # 2023-04-20
            r'(\d{2})-(\d{2})-(\d{4})',  # 01-01-2024
            r'(\d{2})\s*([a-zéû]+)\s*(\d{4})'  # 20 avril 2023
        ]
        
        for pattern in date_format_patterns:
            date_match = re.search(pattern, date_str, re.IGNORECASE)
            if date_match:
                # Extraction du mois et du jour selon le format
                if pattern == r'(\d{4})-(\d{2})-(\d{2})':
                    mois = int(date_match.group(2))
                    jour = int(date_match.group(3))
                elif pattern == r'(\d{2})-(\d{2})-(\d{4})':
                    mois = int(date_match.group(2))
                    jour = int(date_match.group(1))
                else:
                    jour = int(date_match.group(1))
                    mois_str = date_match.group(2)
                    mois = mois_mapping.get(mois_str.lower())
                
                return [mois, jour]
        
        # Formats textuels
        jour_match = re.search(r'(\d+)(?:er)?', date_str)
        mois_match = re.search(r'\b(' + '|'.join(mois_mapping.keys()) + r')\b', date_str)
        
        if not jour_match or not mois_match:
            raise ValueError(f"Format de date invalide : {date_str}")
        
        jour = int(jour_match.group(1))
        mois = mois_mapping[mois_match.group(1)]
        
        return [mois, jour]
    
    except Exception as e:
        print(f"Erreur LLM : Impossible de parser la date via LLM : {date_str}. Utilisation de la méthode de secours.")
        raise

def parse_prompt_with_llm(client, prompt):
    try:
        # Prompt pour l'analyse intelligente
        llm_prompt = f"""
        Tu es un assistant spécialisé dans l'analyse de prompts de gestion de projet.
        Analyse le prompt suivant et réponds au format JSON avec précision :

        Règles :
        - Si le prompt concerne la création d'un projet, fournis :
          * 'action': 'create'
          * 'task_name': nom du projet
          * 'start_date': date de début
          * 'end_date': date de fin
          * 'color': couleur du projet (optionnel)

        - Si le prompt concerne la suppression d'un projet, fournis :
          * 'action': 'delete'
          * 'task_name': nom du projet à supprimer

        - Si le prompt concerne la modification d'un projet, fournis :
          * 'action': 'update'
          * 'task_name': nom du projet à modifier
          * 'start_date': nouvelle date de début (optionnel)
          * 'end_date': nouvelle date de fin (optionnel)
          * 'color': nouvelle couleur (optionnel)

        - Si le prompt concerne la réorganisation verticale des projets, fournis :
          * 'action': 'reorder'
          * 'task_name': nom du projet à repositionner
          * 'second_project': nom du projet de référence (optionnel)
          * 'position': 'first' ou 'before' selon la demande

        Prompt à analyser : {prompt}
        """
        
        # Appel à l'API Ollama
        response = client.chat(
            model='llama3.2',
            messages=[
                {"role": "system", "content": "Tu es un assistant spécialisé dans l'analyse précise de prompts de gestion de projet."},
                {"role": "user", "content": llm_prompt}
            ]
        )
        
        # Vérification de la réponse
        message_content = response.get('message', {}).get('content', '')
        if not message_content:
            raise ValueError("Aucune réponse valide du LLM")
        
        # Extraction du JSON à l'intérieur des backticks
        json_match = re.search(r'```json\n(.*?)\n```', message_content, re.DOTALL)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = message_content
        
        # Tentative de parsing JSON
        try:
            task_info = json.loads(json_str)
        except json.JSONDecodeError:
            # Log détaillé de l'erreur de parsing JSON
            print(f"Erreur de parsing JSON. Contenu reçu : {json_str}")
            raise ValueError(f"Impossible de parser la réponse JSON du LLM : {json_str}")
        
        # Traitement en fonction de l'action
        if task_info.get('action') == 'reorder':
            # Extraction du nom du projet à repositionner
            reorder_match = re.search(r"(?:je veux que|placer|positionner)\s*(?:le)?\s*projet\s*['\"]([^'\"]+)['\"]\s*(?:soit|être)\s*(?:avant|premier|en premier|en haut)", prompt, re.IGNORECASE)
            
            if reorder_match:
                first_project = reorder_match.group(1)
            elif task_info.get('task_name'):
                first_project = task_info['task_name']
            else:
                raise ValueError("Impossible d'extraire le nom du projet à repositionner")
            
            reorder_info = {
                'action': 'reorder',
                'first_project': first_project
            }
            
            # Gestion optionnelle du projet de référence
            if 'second_project' in task_info and task_info['second_project']:
                reorder_info['second_project'] = task_info['second_project']
            
            return reorder_info
        
        # Autres actions (création, suppression, modification)
        elif task_info.get('action') == 'delete':
            return {
                'action': 'delete',
                'task_name': task_info['task_name']
            }
        
        elif task_info.get('action') == 'create':
            # Vérification des champs obligatoires
            if not all(key in task_info for key in ['task_name', 'start_date', 'end_date']):
                raise ValueError("Informations de projet incomplètes")
            
            # Convertir les dates en positions de mois
            start_month = parse_date_to_month_position(client, task_info['start_date'])[0]
            end_month = parse_date_to_month_position(client, task_info['end_date'])[0]
            
            # Mapping des couleurs
            color_map = {
                'rouge': [255, 0, 0],
                'bleu': [0, 0, 255],
                'vert': [0, 255, 0],
                'jaune': [255, 255, 0],
                'orange': [255, 165, 0],
                'violet': [128, 0, 128],
                'rose': [255, 192, 203],
                'marron': [165, 42, 42]
            }
            
            return {
                'action': 'create',
                'task_name': task_info['task_name'],
                'start_month': start_month,
                'end_month': end_month,
                'color_rgb': color_map.get(task_info.get('color', 'bleu').lower(), [0, 0, 255])
            }
        
        elif task_info.get('action') == 'update':
            # Vérification des champs
            if 'task_name' not in task_info:
                raise ValueError("Nom du projet manquant pour la modification")
            
            update_info = {
                'action': 'update',
                'task_name': task_info['task_name']
            }
            
            # Gestion optionnelle des dates et couleur
            if 'start_date' in task_info:
                update_info['start_month'] = parse_date_to_month_position(client, task_info['start_date'])[0]
            
            if 'end_date' in task_info:
                update_info['end_month'] = parse_date_to_month_position(client, task_info['end_date'])[0]
            
            if 'color' in task_info:
                color_map = {
                    'rouge': [255, 0, 0],
                    'bleu': [0, 0, 255],
                    'vert': [0, 255, 0],
                    'jaune': [255, 255, 0],
                    'orange': [255, 165, 0],
                    'violet': [128, 0, 128],
                    'rose': [255, 192, 203],
                    'marron': [165, 42, 42]
                }
                update_info['color_rgb'] = color_map.get(task_info['color'].lower(), [0, 0, 255])
            
            return update_info
        
        else:
            raise ValueError(f"Action non reconnue : {task_info.get('action')}")
    
    except Exception as e:
        # Log détaillé de l'erreur
        print(f"Erreur lors de l'analyse du prompt par l'IA : {e}")
        print(f"Prompt original : {prompt}")
        
        # Relève l'exception pour arrêter le programme
        raise

def create_task_shape(slide, task_name, start_month, end_month, color_rgb, y_position=None, slide_width=None):
    """Crée une forme de tâche dans le slide"""
    print(f"\n--- Création d'une nouvelle tâche ---")
    print(f"Nom de la tâche : {task_name}")
    print(f"Mois de début : {start_month}")
    print(f"Mois de fin : {end_month}")
    print(f"Couleur RGB : {color_rgb}")
    
    # Utiliser la largeur de slide passée en paramètre ou une valeur par défaut
    if slide_width is None:
        slide_width = Inches(13.33)  # Valeur par défaut 16:9
    
    # Marges
    margin_left = Inches(0.5)
    margin_right = Inches(0.5)
    
    # Largeur effective pour la roadmap
    effective_width = slide_width - (margin_left + margin_right)
    
    # Calculer la largeur d'un mois
    month_width = effective_width / 12
    
    # Ajuster la position de départ en fonction de la position dans le mois
    left_start = margin_left + ((start_month) * month_width)
    left_end = margin_left + ((end_month) * month_width)
    
    # Largeur de la tâche
    width = left_end - left_start
    
    # Analyser les formes existantes
    existing_shapes = [shape for shape in slide.shapes if shape.has_text_frame and shape.text_frame.text != "ROADMAP"]
    print(f"\nNombre de formes existantes : {len(existing_shapes)}")
    

    
    # Si aucune position Y n'est spécifiée, trouver la prochaine position disponible
    if y_position is None:
        y_position = 2.5  # Position par défaut
        
        if existing_shapes:
            # Trouver la position Y la plus basse
            y_position = max(
                shape.top / 914400 + shape.height / 914400 + 0.2 
                for shape in existing_shapes
            )

    
    # Créer la forme
    shape = slide.shapes.add_shape(
        1,  # Rectangle
        left_start,
        Inches(y_position),
        width,
        Inches(0.4)  # Hauteur fixe
    )
    
    # Appliquer le style
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color_rgb)
    
    # Ajouter le texte
    shape.text_frame.text = task_name
    shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    shape.text_frame.paragraphs[0].font.size = Pt(10)
    
    print(f"Tâche '{task_name}' créée avec succès !\n")
    
    return shape

def add_month_grid(slide, slide_width, slide_height):
    # Marges internes
    margin_left = Inches(0.5)
    margin_right = Inches(0.5)
    
    # Calculer la largeur effective
    effective_width = slide_width - (margin_left + margin_right)
    
    # Créer la grille des mois
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

def create_roadmap_slide(prs, task_info):
    """
    Crée ou met à jour un slide de roadmap en fonction des informations de tâche.
    
    Args:
        prs (Presentation): La présentation PowerPoint
        task_info (dict): Informations sur la tâche à créer/modifier/supprimer
    """
    print("\n=== Création/Mise à jour du slide de roadmap ===")
    
    # Vérifier si un slide de roadmap existe déjà, sinon en créer un
    roadmap_slide = None
    for slide in prs.slides:
        if slide.slide_layout.name == "Blank":
            roadmap_slide = slide
            break
    
    if roadmap_slide is None:
        # Créer un nouveau slide si aucun n'existe
        blank_slide_layout = prs.slide_layouts[6]  # Layout vide
        roadmap_slide = prs.slides.add_slide(blank_slide_layout)
        print("Création d'un nouveau slide")
    
    # Dimensions de la slide
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    print(f"Dimensions de la slide : {slide_width/914400:.2f} x {slide_height/914400:.2f} inches")
    
    # Ajouter la grille des mois
    add_month_grid(roadmap_slide, slide_width, slide_height)
    
    # Récupérer toutes les formes de tâches existantes
    task_shapes = [
        shape for shape in roadmap_slide.shapes 
        if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
    ]
    
    # Fonction pour repositionner verticalement les tâches
    def reposition_tasks(shapes):
        # Trier les formes de tâches par position Y croissante
        shapes.sort(key=lambda x: x.top)
        
        # Position Y initiale pour les tâches
        start_y = Inches(2.5)
        task_height = Inches(0.6)  # Hauteur standard d'une tâche
        spacing = Inches(0.2)  # Espacement minimal entre les tâches
        
        for shape in shapes:
            # Repositionner la forme de tâche
            shape.top = start_y
            start_y += task_height + spacing
    
    # Traitement de l'action
    if task_info['action'] == 'create':
        # Création d'une nouvelle tâche
        print("--- Création d'une nouvelle tâche ---")
        task_name = task_info['task_name']
        start_month = task_info['start_month']
        end_month = task_info['end_month']
        color_rgb = task_info['color_rgb']
        
        # Créer la forme de la tâche
        create_task_shape(
            roadmap_slide, 
            task_name, 
            start_month, 
            end_month, 
            color_rgb, 
            y_position=None, 
            slide_width=slide_width
        )
        
        # Récupérer à jour les formes de tâches
        task_shapes = [
            shape for shape in roadmap_slide.shapes 
            if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
        ]
        
        # Repositionner toutes les tâches
        reposition_tasks(task_shapes)
        
        print(f"Tâche '{task_name}' créée avec succès !")
        print(f"Nom de la tâche : {task_name}")
        print(f"Mois de début : {start_month}")
        print(f"Mois de fin : {end_month}")
        print(f"Couleur RGB : {color_rgb}")
    
    elif task_info['action'] == 'delete':
        # Suppression d'une tâche
        task_name = task_info['task_name']
        print(f"Suppression du projet : {task_name}")
        
        # Supprimer la tâche spécifique
        task_deleted = False
        for shape in task_shapes:
            if task_name in shape.text_frame.text:
                sp = shape._element
                sp.getparent().remove(sp)
                print(f"Tâche '{task_name}' supprimée avec succès !")
                task_deleted = True
                break
        
        if task_deleted:
            # Récupérer à jour les formes de tâches
            task_shapes = [
                shape for shape in roadmap_slide.shapes 
                if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
            ]
            
            # Repositionner toutes les tâches
            reposition_tasks(task_shapes)
    
    elif task_info['action'] == 'update':
        # Mise à jour d'une tâche existante
        task_name = task_info['task_name']
        print(f"Mise à jour du projet : {task_name}")
        
        # Trouver la tâche à modifier
        task_shape = None
        for shape in task_shapes:
            if task_name in shape.text_frame.text:
                task_shape = shape
                break
        
        if task_shape is None:
            print(f"Erreur : Projet '{task_name}' non trouvé")
            return roadmap_slide
        
        # Récupérer les informations actuelles
        current_start_month = None
        current_end_month = None
        current_color_rgb = None
        
        # Extraire les informations de la forme existante
        for shape in task_shapes:
            if task_name in shape.text_frame.text:
                # Utiliser la position et la largeur pour déterminer les mois
                slide_width = prs.slide_width
                margin_left = Inches(0.5)
                margin_right = Inches(0.5)
                effective_width = slide_width - (margin_left + margin_right)
                month_width = effective_width / 12
                
                current_start_month = int((shape.left - margin_left) / month_width)
                current_end_month = int((shape.left + shape.width - margin_left) / month_width)
                
                # Correction de l'extraction de la couleur
                try:
                    current_color_rgb = [
                        shape.fill.fore_color.rgb.red, 
                        shape.fill.fore_color.rgb.green, 
                        shape.fill.fore_color.rgb.blue
                    ]
                except AttributeError:
                    # Couleur par défaut si l'extraction échoue
                    current_color_rgb = [0, 0, 255]  # Bleu par défaut
                
                break
        
        # Mettre à jour les informations si spécifiées
        start_month = task_info.get('start_month', current_start_month)
        end_month = task_info.get('end_month', current_end_month)
        color_rgb = task_info.get('color_rgb', current_color_rgb)
        
        # Supprimer l'ancienne forme
        sp = task_shape._element
        sp.getparent().remove(sp)
        
        # Créer une nouvelle forme avec les informations mises à jour
        create_task_shape(
            roadmap_slide, 
            task_name, 
            start_month, 
            end_month, 
            color_rgb, 
            y_position=None, 
            slide_width=slide_width
        )
        
        # Récupérer à jour les formes de tâches
        task_shapes = [
            shape for shape in roadmap_slide.shapes 
            if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
        ]
        
        # Repositionner toutes les tâches
        reposition_tasks(task_shapes)
        
        print(f"Projet '{task_name}' mis à jour avec succès !")
        print(f"Nouveau mois de début : {start_month}")
        print(f"Nouveau mois de fin : {end_month}")
        print(f"Nouvelle couleur RGB : {color_rgb}")
    
    elif task_info['action'] == 'reorder':
        # Réorganisation verticale des projets
        first_project = task_info['first_project']
        second_project = task_info.get('second_project')
        
        print(f"Réorganisation verticale : Projet '{first_project}' en premier")
        
        # Récupérer toutes les formes de tâches
        task_shapes = [
            shape for shape in roadmap_slide.shapes 
            if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
        ]
        
        # Trouver les formes des projets spécifiés
        first_shape = None
        second_shape = None
        
        for shape in task_shapes:
            if first_project in shape.text_frame.text:
                first_shape = shape
            if second_project and second_project in shape.text_frame.text:
                second_shape = shape
        
        if first_shape is None:
            print(f"Erreur : Projet '{first_project}' non trouvé")
            return roadmap_slide
        
        # Si second_project est None, on place first_project en premier
        # Sinon, on place first_project avant second_project
        
        # Retirer first_shape de la liste
        task_shapes.remove(first_shape)
        
        if second_project:
            if second_shape is None:
                print(f"Erreur : Projet '{second_project}' non trouvé")
                return roadmap_slide
            
            # Trouver l'index de second_shape
            second_index = task_shapes.index(second_shape)
            
            # Insérer first_shape avant second_shape
            task_shapes.insert(second_index, first_shape)
        else:
            # Placer first_shape au début
            task_shapes.insert(0, first_shape)
        
        # Repositionner toutes les tâches
        def reposition_tasks(shapes):
            # Trier les formes de tâches par position Y croissante
            start_y = Inches(2.5)
            task_height = Inches(0.6)  # Hauteur standard d'une tâche
            spacing = Inches(0.2)  # Espacement minimal entre les tâches
            
            for shape in shapes:
                # Repositionner la forme de tâche
                shape.top = start_y
                start_y += task_height + spacing
        
        reposition_tasks(task_shapes)
        
        print(f"Projet '{first_project}' repositionné avec succès !")
        
    return roadmap_slide

def main():
    # Charger les dossiers de sortie
    templates_dir = "templates"
    output_dir = "generated"
    
    # Créer les dossiers s'ils n'existent pas
    os.makedirs(templates_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # Chemins complets
    template_path = os.path.join(templates_dir, "roadmap_template.pptx")
    output_path = os.path.join(output_dir, "roadmap.pptx")
    
    # Configuration du client Ollama
    client = Client(host='http://localhost:11434')
    
    # Configuration
    config = load_config()
    
    # Lire les prompts
    prompts = []
    with open('prompt.txt', 'r') as f:
        prompts = f.readlines()
    
    # Charger ou créer la présentation
    if os.path.exists(template_path):
        prs = Presentation(template_path)
        print(f"Chargement du template depuis {template_path}")
    else:
        prs = Presentation()
        print("Création d'une nouvelle présentation")
        
        # Sauvegarde du template vide
        prs.save(template_path)
        print(f"Template vide sauvegardé dans {template_path}")
    
    # Traitement des prompts
    for prompt in prompts:
        prompt = prompt.strip()
        if not prompt:
            continue
        
        try:
            task_info = parse_prompt_with_llm(client, prompt)
            
            if task_info:
                # Créer ou mettre à jour le slide de roadmap
                create_roadmap_slide(prs, task_info)
        
        except Exception as e:
            print(f"Erreur lors du traitement du prompt '{prompt}' : {e}")
    
    # Sauvegarder la présentation une seule fois après tous les traitements
    prs.save(output_path)
    print(f"Présentation sauvegardée dans {output_path}")

if __name__ == "__main__":
    main()