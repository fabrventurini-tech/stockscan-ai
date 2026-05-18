# StockScan AI — Guida all'installazione

Applicazione mobile-first per l'inventario ottico di magazzino ricambi bici.
Usa la fotocamera del dispositivo e l'AI di Anthropic per contare i pezzi.

---

## Prerequisiti

- **Node.js** versione 18 o superiore → https://nodejs.org
- Una **API key Anthropic** → https://console.anthropic.com
- Un account **Vercel** (gratuito) per il deploy → https://vercel.com

---

## 1. Installazione locale (per sviluppo e test)

```bash
# 1. Entra nella cartella del progetto
cd stockscan

# 2. Installa le dipendenze
npm install

# 3. Crea il file con la API key
cp .env.example .env
# Apri .env con un editor di testo e incolla la tua API key Anthropic:
# ANTHROPIC_API_KEY=sk-ant-api03-...

# 4. Avvia il server di sviluppo
npm run dev
```

Apri il browser su **http://localhost:5173**

---

## 2. Deploy su Vercel (accesso da smartphone, uso aziendale)

### Passo 1 — Carica il progetto su GitHub

```bash
# Dalla cartella stockscan
git init
git add .
git commit -m "StockScan AI v1.0"
```

Vai su https://github.com → New repository → crea repo → segui le istruzioni per il push.

### Passo 2 — Collega Vercel a GitHub

1. Vai su https://vercel.com e accedi con GitHub
2. Clicca **Add New Project**
3. Seleziona il repository `stockscan`
4. Framework preset: **Vite**
5. Clicca **Deploy** (il primo deploy parte senza API key — normale)

### Passo 3 — Inserisci la API key su Vercel

1. Nel progetto Vercel → **Settings** → **Environment Variables**
2. Aggiungi:
   - Name: `ANTHROPIC_API_KEY`
   - Value: la tua chiave `sk-ant-api03-...`
   - Environments: ✅ Production ✅ Preview
3. Clicca **Save**
4. Vai su **Deployments** → clicca i tre puntini sull'ultimo deploy → **Redeploy**

Vercel ti assegna un URL tipo: `https://stockscan-ai.vercel.app`

---

## 3. Installazione come app su smartphone (PWA)

### Android (Chrome)
1. Apri l'URL Vercel in Chrome
2. Tocca il menu (⋮ in alto a destra)
3. Tocca **"Aggiungi a schermata Home"**
4. Conferma → l'icona StockScan appare nella home

### iPhone (Safari)
1. Apri l'URL in Safari
2. Tocca il tasto **Condividi** (quadrato con freccia su)
3. Scorri e tocca **"Aggiungi a schermata Home"**
4. Conferma

> ⚠️ Su iPhone il BarcodeDetector non è supportato da Safari.
> Lo scanner barcode funziona su Android con Chrome.

---

## 4. Struttura del progetto

```
stockscan/
├── api/
│   └── chat.js          ← Proxy serverless (nasconde la API key)
├── public/
│   └── manifest.json    ← Configurazione PWA
├── src/
│   ├── main.jsx         ← Entry point React
│   └── App.jsx          ← Applicazione principale
├── .env.example         ← Template variabili d'ambiente
├── .env                 ← API key (NON caricare su GitHub)
├── .gitignore
├── index.html
├── package.json
└── vite.config.js
```

---

## 5. Aggiornamenti futuri

Per aggiornare l'app basta modificare i file e fare push su GitHub:
```bash
git add .
git commit -m "descrizione modifica"
git push
```
Vercel rileva il push e rideploya in automatico in ~1 minuto.

---

## 6. Integrazione Ad Hoc Revolution Web (prossimo step)

L'export Excel attuale è pensato per l'importazione manuale.
La fase successiva prevede una chiamata REST alle API Zucchetti per
aggiornare le giacenze direttamente dal gestionale al termine dell'inventario.
