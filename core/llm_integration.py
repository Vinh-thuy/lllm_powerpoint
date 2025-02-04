from openai import OpenAI
import yaml

with open('../config.yaml') as f:
    config = yaml.safe_load(f)

client = OpenAI(api_key=config['openai_key'])

PROMPT_TEMPLATE = '''
Analyse ce prompt utilisateur et extrait les modifications à apporter :

Prompt: {user_input}

Format de sortie YAML :
slides:
  - id: 0
    updates:
      title: "Nouveau titre"
  - id: 1
    updates:
      dates:
        - "2024-03-01 → 2024-03-05"
        - "2024-04-10 → 2024-04-12"
'''

def process_prompt(user_input):
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{
            "role": "user",
            "content": PROMPT_TEMPLATE.format(user_input=user_input)
        }],
        temperature=0.3
    )
    return yaml.safe_load(response.choices[0].message.content)
