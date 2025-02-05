import sqlite3
import json
from datetime import datetime
import re

def normalize_text(text):
    """
    Normalise le texte en supprimant les caractères spéciaux et en uniformisant les espaces.
    
    Args:
        text (str): Texte à normaliser
    
    Returns:
        str: Texte normalisé
    """
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)  # Remplacer les espaces multiples par un seul
    text = re.sub(r'[^\w\s]', '', text)  # Supprimer la ponctuation
    return text.lower()

class TaskDatabase:
    def __init__(self, db_path='tasks.db'):
        """
        Initialise la connexion à la base de données SQLite.
        
        Args:
            db_path (str): Chemin vers le fichier de base de données
        """
        self.db_path = db_path
        self._create_table()
    
    def _create_table(self):
        """
        Crée la table des tâches si elle n'existe pas.
        Vérifie et ajoute les colonnes start_date, end_date et raw_prompt si nécessaire.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Créer la table de base
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_name TEXT NOT NULL,
                    start_month INTEGER,
                    start_position REAL,
                    end_month INTEGER,
                    end_position REAL,
                    color_rgb TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Vérifier et ajouter les colonnes start_date, end_date et raw_prompt
            columns_to_add = [
                ('start_date', 'DATETIME'),
                ('end_date', 'DATETIME'),
                ('raw_prompt', 'TEXT')
            ]
            
            for column_name, column_type in columns_to_add:
                try:
                    # Essayer de sélectionner la colonne
                    cursor.execute(f'SELECT {column_name} FROM tasks LIMIT 1')
                except sqlite3.OperationalError:
                    # La colonne n'existe pas, l'ajouter
                    cursor.execute(f'ALTER TABLE tasks ADD COLUMN {column_name} {column_type}')
            
            conn.commit()
    
    def insert_task(self, task_info):
        """
        Insère une nouvelle tâche dans la base de données.
        
        Args:
            task_info (dict): Informations de la tâche parsées
        
        Returns:
            int: ID de la tâche insérée
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Préparer les valeurs
            task_name = normalize_text(task_info.get('task_name', 'Unnamed Task'))
            
            # Extraire start_month et end_month
            start_month = task_info.get('start_month', [None, None])
            end_month = task_info.get('end_month', [None, None])
            
            # Extraire start_date et end_date
            start_date = task_info.get('start_date', None)
            end_date = task_info.get('end_date', None)
            
            # Convertir color_rgb en chaîne JSON si nécessaire
            color_rgb = (json.dumps(task_info['color_rgb']) 
                         if 'color_rgb' in task_info and task_info['color_rgb'] is not None 
                         else None)
            
            cursor.execute('''
                INSERT INTO tasks (
                    task_name, 
                    start_month, start_position, 
                    end_month, end_position, 
                    color_rgb,
                    start_date,
                    end_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                task_name,
                start_month[0] if start_month[0] is not None else None, 
                start_month[1] if start_month[1] is not None else None,
                end_month[0] if end_month[0] is not None else None, 
                end_month[1] if end_month[1] is not None else None,
                color_rgb,
                start_date,
                end_date
            ))
            conn.commit()
            
            return cursor.lastrowid
    
    def upsert_task(self, task_info, raw_prompt=None):
        """
        Insère une nouvelle tâche ou met à jour une tâche existante basée sur le task_name.
        
        Args:
            task_info (dict): Informations de la tâche parsées
            raw_prompt (str, optional): Texte brut du prompt
        
        Returns:
            int: ID de la tâche insérée ou mise à jour
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Préparer les valeurs
            task_name = normalize_text(task_info.get('task_name', 'Unnamed Task'))
            
            # Convertir color_rgb en chaîne JSON si nécessaire
            color_rgb = (json.dumps(task_info['color_rgb']) 
                         if 'color_rgb' in task_info and task_info['color_rgb'] is not None 
                         else None)
            
            # Vérifier si la tâche existe déjà
            cursor.execute('SELECT * FROM tasks WHERE task_name = ?', (task_name,))
            existing_task = cursor.fetchone()
            
            if existing_task:
                # Préparer les colonnes à mettre à jour
                update_fields = {}
                
                # Colonnes possibles à mettre à jour
                columns_mapping = {
                    'start_month': task_info.get('start_month', [None, None])[0],
                    'start_position': task_info.get('start_month', [None, None])[1],
                    'end_month': task_info.get('end_month', [None, None])[0],
                    'end_position': task_info.get('end_month', [None, None])[1],
                    'color_rgb': color_rgb,
                    'start_date': task_info.get('start_date'),
                    'end_date': task_info.get('end_date'),
                    'raw_prompt': raw_prompt
                }
                
                # Ne conserver que les valeurs explicitement fournies et non-None
                keys_to_check = ['start_month', 'start_position', 'end_month', 'end_position', 
                                 'color_rgb', 'start_date', 'end_date', 'raw_prompt']
                
                update_fields = {
                    k: columns_mapping[k] 
                    for k in keys_to_check 
                    if k in task_info and columns_mapping[k] is not None
                }
                
                if update_fields:
                    # Construire la requête de mise à jour
                    set_clause = ", ".join([f"{col} = ?" for col in update_fields.keys()])
                    update_query = f"""
                        UPDATE tasks 
                        SET {set_clause}, created_at = CURRENT_TIMESTAMP
                        WHERE task_name = ?
                    """
                    
                    # Préparer les valeurs
                    update_values = list(update_fields.values()) + [task_name]
                    
                    cursor.execute(update_query, update_values)
                    task_id = existing_task[0]
                else:
                    # Aucune mise à jour n'est nécessaire
                    task_id = existing_task[0]
            
            else:
                # Insérer une nouvelle tâche
                start_month = task_info.get('start_month', [None, None])
                end_month = task_info.get('end_month', [None, None])
                
                # Extraire start_date et end_date
                start_date = task_info.get('start_date', None)
                end_date = task_info.get('end_date', None)
                
                cursor.execute('''
                    INSERT INTO tasks (
                        task_name, 
                        start_month, start_position, 
                        end_month, end_position, 
                        color_rgb,
                        start_date,
                        end_date,
                        raw_prompt
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    task_name,
                    start_month[0] if start_month[0] is not None else None, 
                    start_month[1] if start_month[1] is not None else None,
                    end_month[0] if end_month[0] is not None else None, 
                    end_month[1] if end_month[1] is not None else None,
                    color_rgb,
                    start_date,
                    end_date,
                    raw_prompt
                ))
                task_id = cursor.lastrowid
            
            conn.commit()
            
            return task_id
    
    def get_task_by_name(self, task_name):
        """
        Récupère une tâche par son nom.
        
        Args:
            task_name (str): Nom de la tâche
        
        Returns:
            dict or None: Informations de la tâche
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM tasks WHERE task_name = ? ORDER BY created_at DESC LIMIT 1', (normalize_text(task_name),))
            row = cursor.fetchone()
            
            return dict(row) if row else None
    
    def list_tasks(self, limit=50):
        """
        Liste les tâches récentes.
        
        Args:
            limit (int): Nombre maximum de tâches à retourner
        
        Returns:
            list: Liste des tâches
        """
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM tasks ORDER BY start_date ASC LIMIT ?', (limit,))
            rows = cursor.fetchall()
            
            return [dict(row) for row in rows]
    
    def delete_task(self, task_info, raw_prompt=None):
        """
        Supprime une tâche de la base de données en utilisant son nom.
        
        Args:
            task_info (dict): Informations de la tâche à supprimer
            raw_prompt (str, optional): Prompt original pour référence
        
        Returns:
            bool: True si la suppression a réussi, False sinon
        """
        # Extraire et normaliser le nom de la tâche
        task_name = normalize_text(task_info.get('task_name'))
        
        # Vérifier que le nom de tâche est présent
        if not task_name:
            print("Erreur : Aucun nom de tâche fourni pour la suppression")
            return False
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Récupérer tous les noms de tâches
                cursor.execute('SELECT task_name FROM tasks')
                existing_tasks = cursor.fetchall()
                
                # Trouver la tâche correspondante après normalisation
                matching_task = None
                for (existing_task_name,) in existing_tasks:
                    if normalize_text(existing_task_name) == task_name:
                        matching_task = existing_task_name
                        break
                
                if matching_task:
                    # Exécuter la suppression avec le nom de tâche original
                    cursor.execute('DELETE FROM tasks WHERE task_name = ?', (matching_task,))
                    
                    # Vérifier si une ligne a été supprimée
                    if cursor.rowcount > 0:
                        print(f"Tâche '{matching_task}' supprimée avec succès")
                        conn.commit()
                        return True
                
                print(f"Aucune tâche trouvée correspondant à '{task_name}'")
                return False
        
        except sqlite3.Error as e:
            print(f"Erreur lors de la suppression de la tâche : {e}")
            return False
