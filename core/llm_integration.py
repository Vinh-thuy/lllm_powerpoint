def process_prompt(prompt):
    # Implémentation basique de génération de roadmap
    return {
        "slides": [
            {
                "id": 0,
                "updates": [
                    {"text": f"Roadmap générée pour : {prompt}"}
                ]
            }
        ]
    }
