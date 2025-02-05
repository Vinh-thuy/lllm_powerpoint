import sqlite3
import json

def convert_db_task_to_task_info(db_task):
    """
    Convertit une tâche de la base de données en task_info.
    
    Args:
        db_task (dict): Tâche récupérée de la base de données
    
    Returns:
        dict: Tâche au format task_info
    """
    # Convertir color_rgb de JSON à liste si nécessaire
    color_rgb = json.loads(db_task['color_rgb']) if db_task['color_rgb'] else None
    
    task_info = {
        "type": "create",  # Par défaut, toujours "create"
        "task_name": db_task['task_name'],
        "start_month": [
            db_task['start_month'] if db_task['start_month'] is not None else None,
            db_task['start_position'] if db_task['start_position'] is not None else None
        ],
        "end_month": [
            db_task['end_month'] if db_task['end_month'] is not None else None,
            db_task['end_position'] if db_task['end_position'] is not None else None
        ],
        "color_rgb": color_rgb
    }
    
    return task_info

def retrieve_tasks_from_sqlite(db_path='tasks.db', limit=None):
    """
    Récupère les tâches de la base de données SQLite.
    
    Args:
        db_path (str): Chemin vers le fichier de base de données
        limit (int, optional): Nombre maximum de tâches à récupérer
    
    Returns:
        list: Liste de tâches au format task_info
    """
    # Connexion à la base de données
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Préparer la requête
    query = 'SELECT * FROM tasks ORDER BY created_at DESC'
    params = []
    
    if limit is not None:
        query += ' LIMIT ?'
        params.append(limit)
    
    # Exécuter la requête
    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    # Convertir les tâches
    task_infos = [convert_db_task_to_task_info(dict(row)) for row in rows]
    
    # Fermer la connexion
    conn.close()
    
    return task_infos

def main():
    # Récupérer toutes les tâches
    all_tasks = retrieve_tasks_from_sqlite()
    print(all_tasks)
    

if __name__ == "__main__":
    main()
