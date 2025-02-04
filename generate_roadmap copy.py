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

def create_task_shape(slide, task_name, start_month, end_month, color_rgb, y_position):
    """Crée une forme de tâche dans le slide"""
    # Dimensions et positions
    month_width = Inches(0.67)  # 8 inches / 12 months
    task_height = Inches(0.4)
    left_margin = Inches(1)
    
    # Calculer la position et la largeur
    left = left_margin + (start_month * month_width)
    width = ((end_month - start_month) + 1) * month_width
    
    # Créer la forme
    shape = slide.shapes.add_shape(
        1,  # Rectangle
        left,
        Inches(y_position),
        width,
        task_height
    )
    
    # Appliquer le style
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color_rgb)
    
    # Ajouter le texte
    shape.text_frame.text = task_name
    shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    shape.text_frame.paragraphs[0].font.size = Pt(10)
    
    return shape

def parse_prompt_with_llm(client, prompt):
    """Analyse le prompt en utilisant GPT-4o-mini"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system", 
                    "content": """Tu es un assistant spécialisé dans l'extraction d'informations de tâches à partir de prompts en français.
Réponds TOUJOURS au format JSON suivant :
{
    "name": "Nom du projet",
    "start_month": 0-11,  # Janvier = 0, Décembre = 11
    "end_month": 0-11,    # Janvier = 0, Décembre = 11
    "color": [R, G, B]    # Valeurs RGB de la couleur
}

Exemples :
1. "creer un projet Deuxs qui demarre le 1er juin et fini le 15 aout" 
→ {"name": "Deuxs", "start_month": 5, "end_month": 7, "color": [0, 0, 255]}

2. "ajoute le projet TOTO se déroule du 1er février au 13 juin (couleur : orange)"
→ {"name": "TOTO", "start_month": 1, "end_month": 5, "color": [255, 165, 0]}

Couleurs possibles :
- rouge : [255, 0, 0]
- vert : [0, 255, 0]
- bleu : [0, 0, 255]
- orange : [255, 165, 0]
- jaune : [255, 255, 0]
- violet : [128, 0, 128]

Si la couleur n'est pas spécifiée, utilise bleu [0, 0, 255]."""},
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            response_format={"type": "json_object"},
            max_tokens=300
        )
        
        # Récupérer et parser la réponse JSON
        task_info = json.loads(response.choices[0].message.content)
        
        # Validation des données
        if (isinstance(task_info.get('start_month'), int) and 
            isinstance(task_info.get('end_month'), int) and 
            0 <= task_info['start_month'] <= 11 and 
            0 <= task_info['end_month'] <= 11 and
            isinstance(task_info.get('name'), str) and 
            isinstance(task_info.get('color'), list) and 
            len(task_info['color']) == 3):
            return task_info
        else:
            print("Format de réponse LLM invalide")
            return None
    
    except Exception as e:
        print(f"Erreur lors de l'analyse du prompt par LLM : {e}")
        return None

def create_roadmap_slide(prs, task_info=None):
    """Crée ou met à jour un slide de roadmap"""
    # Utiliser le premier slide ou en créer un nouveau
    if len(prs.slides) > 0:
        slide = prs.slides[0]
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Layout vide
    
    # # Titre "ROADMAP"
    # if not any(shape.has_text_frame and shape.text_frame.text == "ROADMAP" for shape in slide.shapes):
    #     title = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.5))
    #     title.text_frame.text = "ROADMAP"
    #     title.text_frame.paragraphs[0].font.size = Pt(24)
    #     title.text_frame.paragraphs[0].font.bold = True
    #     title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
    
    # # Grille des mois (si elle n'existe pas déjà)
    # if not any(shape.has_table for shape in slide.shapes):
    #     months_box = slide.shapes.add_table(2, 12, Inches(1), Inches(1.5), Inches(8), Inches(0.5)).table
    #     months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        
    #     for i, month in enumerate(months):
    #         cell = months_box.cell(0, i)
    #         cell.text = month
    #         cell.text_frame.paragraphs[0].font.size = Pt(8)
    #         cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Ajouter la nouvelle tâche si spécifiée
    if task_info:
        # Trouver la prochaine position Y disponible
        y_position = 2.5
        for shape in slide.shapes:
            if shape.has_text_frame and not shape.text_frame.text == "ROADMAP":
                y_position = max(y_position, shape.top / 914400 + shape.height / 914400 + 0.2)  # Conversion EMU to inches
        
        create_task_shape(
            slide,
            task_info['name'],
            task_info['start_month'],
            task_info['end_month'],
            task_info['color'],
            y_position
        )
    
    return slide

def main():
    # Créer le dossier de sortie
    output_dir = "generated"
    output_path = os.path.join(output_dir, "roadmap.pptx")
    
    # Charger ou créer la présentation
    if os.path.exists(output_path):
        prs = Presentation(output_path)
    else:
        prs = Presentation()
        # Définir la taille des slides en 16:9
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
    
    # Récupérer la clé API
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        print("Erreur : La clé API OpenAI n'est pas définie dans le fichier .env")
        return
    
    client = OpenAI(api_key=api_key)
    
    # Traiter le prompt
    prompt = input("Entrez votre demande (ex: ajoute le projet TOTO se déroule du 1er février au 13 juin (couleur : orange)) : ")
    task_info = parse_prompt_with_llm(client, prompt)
    
    if task_info:
        # Créer ou mettre à jour le slide
        create_roadmap_slide(prs, task_info)
        
        # Sauvegarder la présentation
        os.makedirs(output_dir, exist_ok=True)
        prs.save(output_path)
        print(f"Présentation mise à jour : {output_path}")
    else:
        print("Impossible d'extraire les informations de tâche du prompt")

if __name__ == "__main__":
    main()
