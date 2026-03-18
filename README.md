# Portfolio AM — Dashboard Streamlit

Dashboard interattiva per il portfolio di impianti rinnovabili Moncada Energy Group.

## 🚀 Deploy su Streamlit Cloud (gratis, 5 minuti)

### 1. Prepara il repository GitHub
```bash
# Crea un repo GitHub e carica questi file:
# - app.py
# - requirements.txt
# - portfolio_integrato_Q1_2026.xlsx  (opzionale — si può caricare dall'app)
```

### 2. Deploy su Streamlit Cloud
1. Vai su **https://share.streamlit.io**
2. Accedi con GitHub
3. Clicca **New app**
4. Seleziona il repository → `app.py` → **Deploy**

In 2-3 minuti la dashboard è online con un URL condivisibile.

---

## 💻 Esecuzione locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

Apri il browser su **http://localhost:8501**

---

## 📂 Aggiornamento dati

1. Aggiorna il file Excel `portfolio_integrato_Q1_2026.xlsx` con i nuovi dati
2. Nell'app clicca **"Carica Excel aggiornato"** nella sidebar
3. La dashboard si ricalcola automaticamente

---

## 📊 Sezioni disponibili

| Tab | Contenuto |
|-----|-----------|
| 📊 Produzione | Energia mensile FV vs Eolico, top impianti |
| 💰 Finanziario | Fatturato netto, EUR/MWh, tabella dettaglio |
| 💹 Incentivi | Acconto vs Consuntivo GSE, delta mensile |
| ⚡ Analisi Perdite | E. teorica vs reale, % perdita per impianto |
| 🏭 Impianti | Anagrafica, scatter En. vs Fatturato |

---

## 🔧 Filtri disponibili (sidebar)

- **Tipo impianto**: Fotovoltaico / Eolico
- **Mese**: Gennaio / Febbraio / Marzo
