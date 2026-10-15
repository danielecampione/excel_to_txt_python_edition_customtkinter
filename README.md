# Excel to TXT - Python Edition CustomTkinter v0.03a

Conversione del primo foglio Excel in TXT con selezione dinamica delle colonne e join opzionale con un documento secondario.

## Funzioni
- file `.xlsx` e `.xls`;
- trascinamento oppure selezione tramite finestra;
- colonne principali configurabili;
- chiavi separate ed esplicite per documento principale e secondario;
- join su chiavi accoppiate;
- fallback opzionale senza l'ultima chiave, solo se univoco;
- warning per duplicati, fallback e ambiguità;
- elaborazione in background;
- output UTF-8 accanto al documento principale;
- tema e animazioni CustomTkinter conservati.

## Avvio
```bash
pip install -r requirements.txt
python main.py
```

Autore: Daniele Campione  
Licenza: GNU General Public License
