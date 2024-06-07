# Python PowerPoint Translate
Automatically translates the whole PowerPoint document to another language **without messing up the styling of the presentation**

## Install
on linux env use

````
sudo apt update
sudo apt install python3
sudo apt install python3-pip

pip install deep-translator python-pptx tqdm
````

## Usage
Simply call the script providing the path to the pptx file you want to translate

````
python3 ppt_translate.py path_to_pptx_file
````

Default source language is `de` and default target is `en`. You can provide the desired languagey by setting `--source` and `--target` parameters:

````
python3 ppt_translate.py path_to_pptx_file --source tr --target ru
````

Additional info is provided with `-h` parameter

## Output
When everything works as planned you will receive a new file with a `_translated.pptx` suffix

## Features
- Translates the whole document
- Translates all text without messing up the styling / design of the document
- displays a progress bar so user can track progress

## Known Issues
- might have trouble translating text inside Smart Arts
