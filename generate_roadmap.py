import os
from dotenv import load_dotenv
from openai import OpenAI
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
        
        response = client.chat.completions.create(
            model="gpt-4o",  # Utilisation du modèle GPT-4o qui supporte la vision
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
        return response.choices[0].message.content
    except Exception as e:
        print(f"Erreur détaillée lors de l'analyse de l'image : {type(e).__name__} - {str(e)}")
        return None

def parse_date_to_month_position(client, date_str):
    """
    Convertit une date en position de mois avec précision.
    
    Utilise l'IA pour analyser la date avec une méthode de secours locale.
    
    Règles de positionnement :
    - 1-10 : début du mois (0.0)
    - 11-20 : milieu du mois (0.5)
    - 21-31 : fin du mois (1.0)
    
    Retourne un tuple (mois, position_dans_mois)
    """
    # Dictionnaire de mapping des mois avec correction de frappe
    month_map = {
        'janvier': 0, 'jan': 0, 'janv': 0, 'janveir': 0, 'janiver': 0,
        'février': 1, 'fevrier': 1, 'feb': 1, 'fev': 1,
        'mars': 2, 'mar': 2,
        'avril': 3, 'avr': 3, 'apr': 3,
        'mai': 4, 'may': 4,
        'juin': 5, 'jun': 5,
        'juillet': 6, 'jul': 6,
        'août': 7, 'aout': 7, 'aug': 7,
        'septembre': 8, 'sep': 8,
        'octobre': 9, 'oct': 9,
        'novembre': 10, 'nov': 10,
        'décembre': 11, 'decembre': 11, 'dec': 11
    }
    
    # Nettoyer et normaliser la chaîne
    date_str = date_str.lower().strip()
    
    try:
        # Tentative d'analyse avec l'API OpenAI
        prompt = f"""
        Analyse la date suivante et réponds uniquement avec le jour et le mois au format "JJ mois".
        Corrige toute erreur de frappe ou interprétation.
        
        Date à analyser : {date_str}
        """
        
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Tu es un assistant spécialisé dans l'analyse de dates en français."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=50,
            temperature=0.2
        )
        
        # Extraction de la réponse
        parsed_date = response.choices[0].message.content.strip().lower()
        
        # Extraction du jour et du mois
        jour_match = re.search(r'\b(\d+)\b', parsed_date)
        mois_match = re.search(r'\b(' + '|'.join(month_map.keys()) + r')\b', parsed_date)
        
        if not (jour_match and mois_match):
            raise ValueError(f"Impossible de parser la date via LLM : {parsed_date}")
        
        jour = int(jour_match.group(1))
        mois = month_map[mois_match.group(1)]
    
    except Exception as e:
        # Méthode de secours locale
        print(f"Erreur LLM : {e}. Utilisation de la méthode de secours.")
        
        # Extraction du jour
        jour_match = re.search(r'\b(\d+)\b', date_str)
        
        # Extraction du mois
        mois_match = None
        for month_key in month_map.keys():
            if month_key in date_str:
                mois_match = month_key
                break
        
        if not (jour_match and mois_match):
            raise ValueError(f"Format de date invalide : {date_str}")
        
        jour = int(jour_match.group(1))
        mois = month_map[mois_match]
    
    # Calculer la position précise dans le mois
    if 1 <= jour <= 10:
        position = 0.0  # début du mois
    elif 11 <= jour <= 20:
        position = 0.5  # milieu du mois
    else:
        position = 1.0  # fin du mois
    
    print(f"Date analysée : {jour} {mois_match} -> Mois {mois}, Position {position}")
    return (mois, position)

def parse_prompt_with_llm(client, prompt):
    """
    Analyse le prompt de manière intelligente en utilisant l'API OpenAI.
    
    Gère la création, suppression et modification de projets.
    
    Args:
        client (OpenAI): Client OpenAI pour l'analyse
        prompt (str): Prompt à analyser
    
    Returns:
        dict: Informations sur l'action à réaliser
    """
    # Prétraitement du prompt
    prompt = prompt.lower().strip()
    
    # Cas de suppression directe
    delete_match = re.search(r"delete\s*projet\s*['\"]([^'\"]+)['\"]", prompt)
    if delete_match:
        return {
            'action': 'delete',
            'task_name': delete_match.group(1)
        }
    
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

    Prompt à analyser : {prompt}
    """
    
    try:
        # Appel à l'API OpenAI
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "Tu es un assistant spécialisé dans l'analyse précise de prompts de gestion de projet."},
                {"role": "user", "content": llm_prompt}
            ],
            max_tokens=150,
            temperature=0.2
        )
        
        # Extraction de la réponse JSON
        task_info = json.loads(response.choices[0].message.content)
        
        # Traitement en fonction de l'action
        if task_info.get('action') == 'delete':
            return {
                'action': 'delete',
                'task_name': task_info['task_name']
            }
        
        elif task_info.get('action') == 'create':
            # Vérification des champs obligatoires
            if not all(key in task_info for key in ['task_name', 'start_date', 'end_date']):
                raise ValueError("Informations de projet incomplètes")
            
            # Convertir les dates en positions de mois
            start_month = parse_date_to_month_position(client, task_info['start_date'])
            end_month = parse_date_to_month_position(client, task_info['end_date'])
            
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
        
        else:
            raise ValueError(f"Action non reconnue : {task_info.get('action')}")
    
    except json.JSONDecodeError:
        # Tentative de parsing manuel si le JSON échoue
        delete_match = re.search(r"supprime(?:r)?\s*(?:le)?\s*projet\s*['\"]([^'\"]+)['\"]", prompt, re.IGNORECASE)
        if delete_match:
            return {
                'action': 'delete',
                'task_name': delete_match.group(1)
            }
        
        create_match = re.search(r"projet\s*['\"]([^'\"]+)['\"]\s*du\s*(\d+\s*\w+)\s*au\s*(\d+\s*\w+)(?:\s*\(couleur\s*:\s*(\w+)\))?", prompt, re.IGNORECASE)
        if create_match:
            task_name = create_match.group(1)
            start_date = create_match.group(2)
            end_date = create_match.group(3)
            color = create_match.group(4) or 'bleu'
            
            start_month = parse_date_to_month_position(client, start_date)
            end_month = parse_date_to_month_position(client, end_date)
            
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
                'task_name': task_name,
                'start_month': start_month,
                'end_month': end_month,
                'color_rgb': color_map.get(color.lower(), [0, 0, 255])
            }
        
        print(f"Erreur : impossible de parser le prompt '{prompt}'")
        raise ValueError(f"Format de prompt non reconnu : {prompt}")
    
    except Exception as e:
        print(f"Erreur lors de l'analyse du prompt par l'IA : {e}")
        raise ValueError(f"Impossible de parser le prompt : {prompt}")

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
    left_start = margin_left + ((start_month[0] + start_month[1]) * month_width)
    left_end = margin_left + ((end_month[0] + end_month[1]) * month_width)
    
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

def create_roadmap_slide(prs, task_info=None):
    """Crée ou met à jour un slide de roadmap"""
    print("\n=== Création/Mise à jour du slide de roadmap ===")
    
    # Utiliser le premier slide ou en créer un nouveau
    if len(prs.slides) > 0:
        slide = prs.slides[0]
        print(f"Utilisation du slide existant (index 0)")
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Layout vide
        print(f"Création d'un nouveau slide")
    
    # Vérifier si la grille des mois existe déjà
    months_grid_exists = any(shape.has_table for shape in slide.shapes)
    
    # Titre "ROADMAP"
    if not any(shape.has_text_frame and shape.text_frame.text == "ROADMAP" for shape in slide.shapes):
        title = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.5))
        title.text_frame.text = "ROADMAP"
        title.text_frame.paragraphs[0].font.size = Pt(24)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)    
    
    if not months_grid_exists:
        print("Ajout de la grille des mois")
        
        # Calculer les dimensions de la slide
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        print(f"Dimensions de la slide : {slide_width / 914400:.2f} x {slide_height / 914400:.2f} inches")
        
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
    else:
        print("Grille des mois déjà existante")
    
    # Gestion des actions
    if task_info is None:
        return slide
    
    # Action de suppression
    if task_info.get('action') == 'delete':
        task_name_to_delete = task_info['task_name']
        print(f"Suppression du projet : {task_name_to_delete}")
        
        # Parcourir et supprimer les formes de tâches
        for shape in list(slide.shapes):
            if shape.has_text_frame and shape.text_frame.text != "ROADMAP":
                # Vérifier si le nom du projet correspond
                if task_name_to_delete == shape.text_frame.text:
                    # Méthode de suppression
                    xml_element = shape.element
                    xml_element.getparent().remove(xml_element)
                    print(f"Projet '{task_name_to_delete}' supprimé")
                    break
        
        # Repositionner les tâches restantes
        task_shapes = [
            shape for shape in slide.shapes 
            if shape.has_text_frame and shape.text_frame.text != "ROADMAP"
        ]
        
        # Trier les formes de tâches par position Y croissante
        task_shapes.sort(key=lambda x: x.top)
        
        # Repositionner uniquement les tâches
        start_y = Inches(2.5)  # Position Y initiale pour les tâches
        for shape in task_shapes:
            # Repositionner la forme de tâche
            shape.top = start_y
            start_y += Inches(0.6)  # Espacement entre les tâches
    
    # Action de création de projet
    elif task_info.get('action') == 'create':
        create_task_shape(
            slide,
            task_info['task_name'],
            task_info['start_month'],
            task_info['end_month'],
            task_info['color_rgb'],
            slide_width=prs.slide_width
        )
    
    return slide

def main():
    # Créer les dossiers de sortie
    templates_dir = "templates"
    output_dir = "generated"
    
    # Créer les dossiers s'ils n'existent pas
    os.makedirs(templates_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # Chemins complets
    template_path = os.path.join(templates_dir, "roadmap_template.pptx")
    output_path = os.path.join(output_dir, "roadmap.pptx")
    
    # Supprimer le fichier existant dans le répertoire de génération
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
            print(f"Fichier existant supprimé : {output_path}")
        except Exception as e:
            print(f"Erreur lors de la suppression du fichier : {e}")
    
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
    
    # Configuration
    config = load_config()
    api_key = os.getenv('OPENAI_API_KEY')
    client = OpenAI(api_key=api_key)
    
    # Lire les prompts
    prompts = []
    
    # Vérifier si des arguments en ligne de commande sont passés
    if len(sys.argv) > 1:
        # Concaténer tous les arguments en ligne de commande
        prompts = [' '.join(sys.argv[1:])]
    
    # Si pas d'arguments en ligne de commande, essayer de lire le fichier
    if not prompts:
        try:
            with open('prompt.txt', 'r') as f:
                prompts = [line.strip() for line in f.readlines() if line.strip()]
        except FileNotFoundError:
            # Si le fichier n'existe pas, demander une saisie interactive
            prompts = [input("Entrez votre demande (ex: ajoute le projet TOTO se déroule du 1er février au 13 juin (couleur : orange)) : ")]
    
    # Traiter chaque prompt successivement
    for prompt in prompts:
        print(f"\n--- Traitement du prompt : {prompt} ---")
        
        try:
            task_info = parse_prompt_with_llm(client, prompt)
            
            if task_info:
                # Créer ou mettre à jour le slide de roadmap
                create_roadmap_slide(prs, task_info)
                
                # Sauvegarder la présentation après chaque prompt
                prs.save(output_path)
                print(f"Présentation mise à jour : {output_path}")
            else:
                print("Impossible de parser le prompt")
        
        except Exception as e:
            print(f"Erreur lors du traitement du prompt '{prompt}' : {e}")
            traceback.print_exc()
    
    print("\nTraitement de tous les prompts terminé.")

if __name__ == "__main__":
    main()