import sqlite3
import json
from datetime import datetime

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
                ('start_date', 'TEXT'),
                ('end_date', 'TEXT'),
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
            task_name = task_info.get('task_name', 'Unnamed Task')
            
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
            task_name = task_info.get('task_name', 'Unnamed Task')
            
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
            
            # Vérifier si la tâche existe déjà
            cursor.execute('SELECT id FROM tasks WHERE task_name = ?', (task_name,))
            existing_task = cursor.fetchone()
            
            if existing_task:
                # Mettre à jour la tâche existante
                cursor.execute('''
                    UPDATE tasks 
                    SET start_month = ?, start_position = ?, 
                        end_month = ?, end_position = ?, 
                        color_rgb = ?,
                        start_date = ?,
                        end_date = ?,
                        raw_prompt = ?,
                        created_at = CURRENT_TIMESTAMP
                    WHERE task_name = ?
                ''', (
                    start_month[0] if start_month[0] is not None else None, 
                    start_month[1] if start_month[1] is not None else None,
                    end_month[0] if end_month[0] is not None else None, 
                    end_month[1] if end_month[1] is not None else None,
                    color_rgb,
                    start_date,
                    end_date,
                    raw_prompt,
                    task_name
                ))
                task_id = existing_task[0]
            else:
                # Insérer une nouvelle tâche
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
            
            cursor.execute('SELECT * FROM tasks WHERE task_name = ? ORDER BY created_at DESC LIMIT 1', (task_name,))
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
            
            cursor.execute('SELECT * FROM tasks ORDER BY created_at DESC LIMIT ?', (limit,))
            rows = cursor.fetchall()
            
            return [dict(row) for row in rows]

def main():
    # Exemple d'utilisation
    db = TaskDatabase()
    
    # Exemple de tâche
    task_info = {
        'task_name': 'Projet Test',
        'start_month': [2, 0.5],
        'end_month': [6, 1.0],
        'color_rgb': [255, 0, 0],
        'start_date': '2022-01-01',
        'end_date': '2022-12-31'
    }
    
    # Insérer une tâche
    task_id = db.insert_task(task_info)
    print(f"Tâche insérée avec l'ID : {task_id}")
    
    # Récupérer la tâche
    task = db.get_task_by_name('Projet Test')
    print("Tâche récupérée :", task)
    
    # Lister les tâches
    tasks = db.list_tasks()
    print("Liste des tâches :", tasks)

if __name__ == "__main__":
    main()
