# 🚀 Hertz Fatture - Sistema Gestionale Web

Sistema completo per la gestione di preventivi, purchase order e fatture Hertz con generazione automatica XML per Danea Easyfatt.

## ✨ Funzionalità

- ✅ **Upload PDF** con drag & drop
- ✅ **Parsing automatico** preventivi e purchase order
- ✅ **Associazione automatica** per pratica Hertz
- ✅ **Generazione XML** compatibile Easyfatt
- ✅ **Download email** automatico da Gmail (IMAP)
- ✅ **Gestione numerazione** HM/HG separata
- ✅ **Dashboard** con statistiche
- ✅ **Backup automatico** Google Drive (opzionale)

## 📋 Formato File Generati

```
Fatt_040_PO_6440115_GZ605WM.xml
```

Struttura: `Fatt_XXX_PO_YYYYYY_TARGA.xml`

XML compatibile con **Danea Easyfatt** (nodo `<EasyfattDocuments>`)

## 🔧 Configurazione Render

### Variabili d'ambiente richieste:

Nessuna! L'app funziona out-of-the-box.

### Variabili d'ambiente opzionali (Google Drive Backup):

- `GOOGLE_DRIVE_CREDENTIALS` - JSON completo delle credenziali service account
- `GOOGLE_DRIVE_FOLDER_ID` - ID della cartella di backup

### Build Command:
```bash
pip install -r requirements.txt
```

### Start Command:
```bash
gunicorn app:app
```

## 📦 Deploy su Render

1. Fork/Clone questo repository
2. Crea un nuovo **Web Service** su Render
3. Collega il repository
4. Imposta:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: Free
5. Deploy! ✅

## 🔐 Backup Google Drive (Opzionale)

Per attivare il backup automatico:

1. Crea un progetto su Google Cloud Console
2. Attiva Google Drive API
3. Crea un Service Account
4. Scarica le credenziali JSON
5. Condividi una cartella Drive con l'email del service account
6. Su Render → Environment → Aggiungi:
   - `GOOGLE_DRIVE_CREDENTIALS` = contenuto JSON completo
   - `GOOGLE_DRIVE_FOLDER_ID` = ID della cartella

Ogni PDF e XML verrà salvato automaticamente su Google Drive!

## 📊 Dati Iniziali

L'app carica automaticamente i dati da `hertz_data_initial.json` al primo avvio:
- 54 fatture generate
- Configurazione numerazione (HM: 39, HG: 15)
- Configurazione email

## 🐛 Troubleshooting

### Build failed - requirements.txt
Verifica che `runtime.txt` contenga `python-3.11.8`

### Email non funziona
1. Usa una **password per app** Gmail (non la password normale)
2. Verifica configurazione IMAP in Gmail

### File non si caricano con drag&drop
Aggiorna il browser o usa Chrome/Edge

### Google Drive non funziona
Verifica che:
- Le credenziali JSON siano complete
- L'email del service account abbia accesso alla cartella
- L'ID cartella sia corretto

## 📱 Accesso

Dopo il deploy, l'app sarà disponibile su:
```
https://nome-servizio.onrender.com
```

## 🔒 Sicurezza

- Nessuna autenticazione richiesta (accesso diretto)
- Non condividere l'URL pubblicamente
- Usa solo per scopi interni

## 🛠️ Stack Tecnologico

- **Backend**: Flask 3.0.0
- **PDF Processing**: pdfplumber 0.10.3
- **XML Generation**: Python xml.etree
- **Email**: imaplib (Gmail IMAP)
- **Hosting**: Render.com
- **Backup**: Google Drive API (opzionale)

## 📄 Licenza

Uso interno - Tutti i diritti riservati

## 👨‍💻 Supporto

Per problemi o domande, contatta l'amministratore.

---

**Versione**: 2.0.0  
**Ultimo aggiornamento**: Febbraio 2026
