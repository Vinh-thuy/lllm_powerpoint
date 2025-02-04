from pptx import Presentation
from pptx.util import Pt
import re

class TemplateProcessor:
    def __init__(self, template_path):
        self.presentation = Presentation(template_path)
        self.slide_layouts = {
            'title': 0,
            'schedule': 1,
            'content': 2
        }

    def update_slide(self, slide_id, updates):
        slide = self.presentation.slides[slide_id]
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    # Remplacement des variables {{var}}
                    for var, value in updates.items():
                        text = re.sub(f'{{{{{var}}}}}', str(value), text)
                    run.text = text

    def save(self, output_path):
        self.presentation.save(output_path)
