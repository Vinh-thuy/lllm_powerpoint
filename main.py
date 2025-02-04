import argparse
from core.template_processor import TemplateProcessor
from core.llm_integration import process_prompt
import yaml
import os

def main():
    parser = argparse.ArgumentParser(description='Automate PowerPoint updates')
    parser.add_argument('prompt', help='User modification prompt')
    parser.add_argument('-o', '--output', required=True, help='Output file name')
    args = parser.parse_args()
    
    with open('config.yaml') as f:
        config = yaml.safe_load(f)
    
    updates = process_prompt(args.prompt)
    processor = TemplateProcessor(config['template_path'])
    
    for slide_update in updates['slides']:
        processor.update_slide(slide_update['id'], slide_update['updates'])
    
    output_path = os.path.join(config['output_dir'], args.output)
    processor.save(output_path)
    print(f"Présentation générée : {output_path}")

if __name__ == "__main__":
    main()
