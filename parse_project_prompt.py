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

def parse_project_prompt(client, prompt, config):
    """
    Analyse un prompt pour en extraire les informations de projet
    
    Args:
        client (ollama.Client): Client Ollama pour générer la réponse
        prompt (str): Prompt à analyser
        config (dict): Configuration de l'application
    
    Returns:
        dict or list: Informations du/des projet(s)
    """
    # Prompt pour l'extraction des informations de projet
    extraction_prompt = f"""
    INSTRUCTIONS STRICTES POUR L'EXTRACTION DE PROJET :

    1. Analyse le prompt suivant : "{prompt}"

    2. FORMAT DE RÉPONSE OBLIGATOIRE :
    - Un seul projet : {"task_name": "NomProjet", "start_month": [index_mois, position], "end_month": [index_mois, position], "color_rgb": [R,G,B]}
    - Plusieurs projets : [{"task_name": "Projet1", ...}, {"task_name": "Projet2", ...}]

    3. RÈGLES DE CONVERSION :
    - Mois : janvier = 0, décembre = 11
    - Position dans le mois : 0.0 = début, 0.5 = milieu, 1.0 = fin
    - Couleur : Conversion RGB obligatoire

    4. CONTRAINTES :
    - JSON STRICTEMENT VALIDE
    - Pas de texte supplémentaire
    - Retourne [] si aucun projet

    5. EXEMPLE COMPLET :
    Prompt: "Projet ALPHA du 15 février au 30 juin en vert"
    Réponse : [{"task_name": "ALPHA", "start_month": [1, 0.5], "end_month": [5, 1.0], "color_rgb": [0, 128, 0]}]

    GÉNÈRE LA RÉPONSE MAINTENANT :
    """

    # Générer la réponse
    response = client.chat(
        model=config.get('OLLAMA_MODEL', 'llama2'),
        messages=[
            {
                'role': 'system',
                'content': 'Assistant spécialisé en extraction précise de projets.'
            },
            {
                'role': 'user',
                'content': extraction_prompt
            }
        ]
    )

    # Extraire le contenu de la réponse
    content = response['message']['content'].strip()

    # Fonction pour nettoyer et parser le JSON
    def clean_and_parse_json(json_str):
        # Supprimer les espaces et les caractères spéciaux avant et après
        json_str = json_str.strip()
        
        # Supprimer les blocs de code markdown si présents
        if json_str.startswith('```json') and json_str.endswith('```'):
            json_str = json_str[7:-3].strip()
        
        # Essayer de parser le JSON
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            # Tentative de correction des erreurs courantes
            try:
                # Remplacer les guillemets droits par des guillemets simples
                json_str = json_str.replace('"', '"').replace('"', '"')
                return json.loads(json_str)
            except:
                return None

    try:
        # Essayer de parser le JSON
        parsed_data = clean_and_parse_json(content)
        
        # Si c'est une liste, retourner le premier projet
        if isinstance(parsed_data, list):
            if parsed_data:
                return parsed_data[0]
            else:
                print("Aucun projet trouvé dans le prompt.")
                return None
        
        # Si c'est un dictionnaire, le retourner directement
        elif isinstance(parsed_data, dict):
            return parsed_data
        
        else:
            print("Format de réponse invalide.")
            return None

    except Exception as e:
        print(f"Erreur inattendue : {e}")
        print(f"Contenu problématique : {content}")
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
        # Configuration de l'application
        config = {
            'OLLAMA_MODEL': 'mervinpraison/llama3.2-3B-instruct-test-2:8b'
        }
        
        # Créer un client Ollama
        client = ollama.Client()
        
        result = parse_project_prompt(client, prompt, config)
        
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
