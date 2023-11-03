# Python PPT Notes Extractor: A 3AM Conconction
## Motivation
In order to avoid having to manually copy paste notes into my e-notebook for courses with slides upto 60 pages. And thankfully because I get access to .pptx
I decided to make a script that will (hopefully) take text from each slide, append it to a .txt file. 
Version 1.2 will hopefully get better at seperating header/body text

## How To Use

Put .pptx file in the same folder as extractorPPT.py and edit input/output file names (Output support for markdown file [.md]). Then run script
You can also try it by the sample presentations given
## ToDo

1. ~New paragraph by new bullet not end of line~
2. ~Bold/alter formatting (may have to go from .txt to something else?)~
3. Automatically turn to PDF
4. Automate all ".PPTX" files in directory, using some linux magic and turning pyton file to an executable script
5. ~Image support~

## Documentation for libraries

[Python-PPTX Dogs](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)
