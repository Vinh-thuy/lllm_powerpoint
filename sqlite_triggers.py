import sqlite3
import os
from datetime import datetime

def create_month_update_trigger(db_path='tasks.db'):
    """
    Crée un trigger pour mettre à jour automatiquement les mois 
    lors de la modification des dates dans la table tasks.
    
    Args:
        db_path (str): Chemin vers le fichier de base de données SQLite
    """
    try:
        # Connexion à la base de données
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            # Supprimer le trigger existant s'il existe
            cursor.execute('''
                DROP TRIGGER IF EXISTS update_months_trigger
            ''')
            
            # Créer le nouveau trigger
            cursor.execute('''
                CREATE TRIGGER update_months_trigger
                AFTER UPDATE OF start_date, end_date ON tasks
                BEGIN
                    UPDATE tasks SET 
                        start_month = CAST(strftime('%m', NEW.start_date) AS INTEGER) - 1,
                        end_month = CAST(strftime('%m', NEW.end_date) AS INTEGER) - 1
                    WHERE id = NEW.id;
                END;
            ''')
            
            print("Trigger 'update_months_trigger' créé avec succès.")
            
            # Vérifier la création du trigger
            cursor.execute("SELECT name FROM sqlite_master WHERE type='trigger' AND name='update_months_trigger'")
            trigger = cursor.fetchone()
            if trigger:
                print("Trigger existe dans la base de données.")
            else:
                print("ERREUR : Le trigger n'a pas été créé.")
    
    except sqlite3.Error as e:
        print(f"Erreur lors de la création du trigger : {e}")

def test_trigger(db_path='tasks.db'):
    """
    Tester le déclenchement du trigger
    
    Args:
        db_path (str): Chemin vers le fichier de base de données SQLite
    """
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            # Insérer une tâche de test
            cursor.execute('''
                INSERT INTO tasks (task_name, start_date, end_date) 
                VALUES (?, ?, ?)
            ''', ('Test Trigger', '2023-07-15', '2023-09-30'))
            task_id = cursor.lastrowid
            
            # Mettre à jour la date de début
            cursor.execute('''
                UPDATE tasks 
                SET start_date = ? 
                WHERE id = ?
            ''', ('2023-08-01', task_id))
            
            # Vérifier les valeurs mises à jour
            cursor.execute('SELECT * FROM tasks WHERE id = ?', (task_id,))
            updated_task = cursor.fetchone()
            
            print("\nRésultats du test :")
            for key in updated_task.keys():
                print(f"{key}: {updated_task[key]}")
    
    except sqlite3.Error as e:
        print(f"Erreur lors du test du trigger : {e}")

def main():
    """
    Fonction principale pour tester la création du trigger
    """
    # Chemin par défaut de la base de données
    db_path = 'tasks.db'
    
    # Créer le trigger
    create_month_update_trigger(db_path)
    
    # Tester le trigger
    #test_trigger(db_path)

if __name__ == "__main__":
    main()
