
# bkppt 

This program is a simple tool that converts text files, image files, etc., into a basic PowerPoint presentation. It was designed to help users focus more on content when creating PowerPoint slides. The program is built using python-pptx as its core library. It includes a feature that loads text files and places their content into the title and body sections of PowerPoint slides.

- B K Choi
- Feb 10, 2025

# Install

pip install git+https://github.com/skyboong/bkppt2.git

# Tutorial

```
from pathlib import Path
import os 


from bkppt2 import bkppt as bp 
dict_list = []

file_names1 = ['manual_pyenv1.txt']
file_dir_text = 'text_manual'
file_dir_image = 'images'
file_pre_fix = "manual"
print(f"* file names1 = {file_names1}")

current_dir = Path(os.getcwd())
path_figure_directory = current_dir/ file_dir_image
for each in file_names1:
    each2 = current_dir / file_dir_text / each 
    temp_list = bp.read_text_file_for_pptx(each2)
    dict_list.extend(temp_list)

bp.create_pppt(prefix=file_pre_fix,
                slides_data=dict_list, 
                slide_width=13.33,
                slide_height=7.5,
                default_textbox_height=5.5, 
                default_textbox_width=10.0,
                dir_figure=path_figure_directory, 
                font_name="JetBrains Mono",
                font_size_title=36,
                font_size_level_0=12,
                font_size_level_1=10,
                font_size_level_2=10,
                font_size_level_3=8,
                space_before=10,
                space_after=5,
                space_before1=2,
                space_after1=1,
                color_bold="dark_red",
                auto_paragraphs_two=True,
                auto_paragraphs_threshold=500,
                auto_figure_position=True,
                auto_figure_vertical=False,
                auto_figure_xp=1,
                auto_figure_yp=6,
                auto_figure_delta=1.2,
                auto_figure_height=1)
```

