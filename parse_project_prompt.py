import json
import re
import ollama

def convert_color_to_rgb(color):
    """
    Convertit un nom de couleur en code RGB
    
    Args:
        color (str): Nom de la couleur
    
    Returns:
        list: Code RGB de la couleur
    """
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

def parse_project_prompt(prompt):
    """
    Analyse un prompt de projet en utilisant un modèle LLM Ollama
    
    Args:
        prompt (str): Prompt décrivant un projet
    
    Returns:
        dict: Informations structurées du projet
    """
    # Mapping des mois avec plusieurs variantes
    mois_reference = {
        0: 'jan|janvier',
        1: 'fev|fevrier|février',
        2: 'mar|mars',
        3: 'avr|avril',
        4: 'mai',
        5: 'jun|juin',
        6: 'jul|juillet',
        7: 'aou|août|aout',
        8: 'sep|septembre',
        9: 'oct|octobre',
        10: 'nov|novembre',
        11: 'dec|decembre|décembre'
    }
    
    # Configuration du prompt pour l'analyse
    system_prompt = """
    Tu es un assistant spécialisé dans l'analyse de prompts de projet.
    Extrait les informations suivantes :
    - Nom du projet
    - Date de début (jour et mois)
    - Date de fin (jour et mois)
    - Couleur du projet
    
    Règles importantes :
    - Le mois de janvier est 0
    - Le mois de décembre est 11
    - Vérifie attentivement le mois de fin
    
    Réponds UNIQUEMENT au format JSON suivant :
    {
        "task_name": "Nom du projet",
        "start_month": [index_mois, position_dans_mois],
        "end_month": [index_mois, position_dans_mois],
        "color_rgb": [R, G, B]
    }
    
    Règles pour la position dans le mois :
    - 0.0 : début du mois (1-10)
    - 0.5 : milieu du mois (11-20)
    - 1.0 : fin du mois (21-31)
    
    Exemples :
    Prompt: "je veux un projet 'P1' du 15 mai au 29 decembre (couleur : vert)"
    Réponse: {
        "task_name": "P1",
        "start_month": [4, 0.5],
        "end_month": [11, 1.0],
        "color_rgb": [0, 255, 0]
    }
    """
    
    try:
        # Appel au modèle Ollama
        response = ollama.chat(
            model='mervinpraison/llama3.2-3B-instruct-test-2:8b',
            messages=[
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': prompt}
            ]
        )
        
        # Extraction du contenu de la réponse
        result = response['message']['content'].strip()
        
        # Tentative de parsing JSON
        try:
            parsed_result = json.loads(result)
            return parsed_result
        except json.JSONDecodeError:
            # Fallback : extraction par regex si JSON invalide
            projet_match = re.search(r"'([^']+)'", prompt)
            debut_match = re.search(r"du (\d+ \w+)", prompt)
            fin_match = re.search(r"au (\d+ \w+)", prompt)
            couleur_match = re.search(r"couleur\s*:\s*(\w+)", prompt, re.IGNORECASE)
            
            # Extraction du jour et du mois
            def extraire_mois_et_position(date_match):
                if not date_match:
                    return [0, 0.5]  # Valeur par défaut
                
                jour_str, mois_str = date_match.group(1).split()
                jour = int(jour_str)
                
                # Trouver le mois le plus proche
                mois_index = trouver_mois_le_plus_proche(mois_str, mois_reference)
                
                # Déterminer la position dans le mois
                if jour <= 10:
                    position = 0.0
                elif jour <= 20:
                    position = 0.5
                else:
                    position = 1.0
                
                return [mois_index, position]
            
            # Couleur
            couleur = couleur_match.group(1).lower() if couleur_match else "bleu"
            
            return {
                "task_name": projet_match.group(1) if projet_match else "Projet sans nom",
                "start_month": extraire_mois_et_position(debut_match),
                "end_month": extraire_mois_et_position(fin_match),
                "color_rgb": convert_color_to_rgb(couleur)
            }
    
    except Exception as e:
        print(f"Erreur lors de l'analyse du prompt : {e}")
        return None

def trouver_mois_le_plus_proche(mois_str, mois_reference):
    for mois_index, mois_pattern in mois_reference.items():
        if re.match(mois_pattern, mois_str, re.IGNORECASE):
            return mois_index
    return 0  # Valeur par défaut

def main():
    # Exemple d'utilisation
    prompt = input("Entrez votre prompt de projet : ")
    
    try:
        result = parse_project_prompt(prompt)
        
        print("\nRésultat :")
        
        # Vérifier si result est un dictionnaire
        if not isinstance(result, dict):
            print("Erreur : Le résultat n'est pas un dictionnaire.")
            return
        
        # Convertir en JSON pour s'assurer de la sérialisation
        try:
            res = json.dumps(result, ensure_ascii=False, indent=2)
            print('res : ', res)
            print('type res : ', type(res))
            
            # Tenter de reparser le JSON pour vérifier sa validité
            parsed_res = json.loads(res)
            print('Parsed res: ', parsed_res)
            print('Parsed res TYPE: ', type(parsed_res))
        except json.JSONDecodeError as e:
            print(f"Erreur de JSON : {e}")
            print("Contenu original : ", result)
    
    except Exception as e:
        print(f"Erreur lors de l'analyse du prompt : {e}")

if __name__ == "__main__":
    main()
