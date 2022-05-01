# Tyranny Mini Translator
This mini project dealt with a translation script of the [Tyranny](https://www.paradoxinteractive.com/games/tyranny/about) game through [google translator](https://translate.google.com.br/).

## Requirements
[Python 3.9+](https://www.python.org/)

## Setup

You will need to install the script dependencies:
```
pip install -r requirements.txt
```
Copy this folders to script path:
```
<game_location>\Tyranny\Data\data
<game_location>\Tyranny\Data\data_vx1
<game_location>\Tyranny\Data\data_vx2
<game_location>\Tyranny\Data\data_vx3
```

Update script lines 13 to 16 with you source and targe language:
```
SOURCE_LOCALE = 'en'
TARGET_LOCALE = 'pt'
TARGET_NAME = 'portuguese'
TARGET_VERBOSE = 'Português (Brasil)'
```

## Use


Run the script for the first time:
```
python translate.py
```
Use google translation to translate xlsx files in temp folder.
Run script again and wait finish.
Copy game data folders to original location.
