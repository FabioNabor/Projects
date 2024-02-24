import os
from pathlib import Path

def configCard(card):
    card = str(card).strip().replace('.','').lstrip('0')
    sizecard = len(card)
    if sizecard == 16:
        return f"{card[0:4]}.{card[4:8]}.{card[8:12]}.{card[12:16]}"
    if sizecard <= 12:
        zeros = 12 - sizecard
        card = f'{'0'*zeros}{card}'
        return f"{card[0:4]}.{card[4:9]}.{card[9:12]}"

def verificyExistFile(namefile):
    diretoryhome = Path.home()
    downloads = diretoryhome / "Downloads"
    files = os.listdir(downloads)
    if namefile in files:
        return True
    else:
        return False
