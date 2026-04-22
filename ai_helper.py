"""
AI Helper v7.1 — Modulo opzionale per generazione premesse con Claude AI.
Funziona SOLO se l'utente ha configurato una API key Anthropic nelle impostazioni.
Senza API key il programma funziona perfettamente con i template standard.
"""

import json

# Costo stimato: ~1-3 centesimi per documento (Sonnet)
# Lista modelli in ordine di preferenza: il più recente prima
MODELLI_PREFERITI = [
    "claude-sonnet-4-5-20250929",
    "claude-sonnet-4-20250514",
    "claude-3-5-sonnet-20241022",
    "claude-3-5-sonnet-20240620",
]
MODEL = MODELLI_PREFERITI[0]  # Default
MAX_TOKENS = 2000


def is_ai_disponibile():
    """Verifica se il modulo AI è utilizzabile (API key presente + libreria installata)."""
    try:
        import database as db
        key = db.get_impostazione("anthropic_api_key")
        if not key or len(key) < 10:
            return False
        # Verifica che la libreria httpx o urllib sia disponibile (non serve anthropic SDK)
        return True
    except Exception:
        return False


def get_api_key():
    import database as db
    return db.get_impostazione("anthropic_api_key", "")


def genera_premesse_ai(tipo_determina, dati_affidamento, dati_extra=None):
    """
    Genera le premesse per una determina usando Claude AI.

    Args:
        tipo_determina: str — tipo di determina
        dati_affidamento: dict — dati dell'affidamento dal DB
        dati_extra: dict — dati aggiuntivi (coperture, fattura, ecc.)

    Returns:
        str — testo delle premesse, oppure None se errore/non disponibile
    """
    api_key = get_api_key()
    if not api_key:
        return None

    # Costruisci il contesto per l'AI
    contesto = _costruisci_contesto(tipo_determina, dati_affidamento, dati_extra)
    prompt = _costruisci_prompt(tipo_determina, contesto)

    try:
        import urllib.request
        import ssl

        payload = json.dumps({
            "model": MODEL,
            "max_tokens": MAX_TOKENS,
            "messages": [{"role": "user", "content": prompt}],
            "system": SYSTEM_PROMPT,
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST",
        )

        # Disabilita verifica SSL solo se necessario (ambienti aziendali con proxy)
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        # Estrai il testo dalla risposta
        if data.get("content") and len(data["content"]) > 0:
            return data["content"][0].get("text", "").strip()
        return None

    except Exception as e:
        print(f"[AI Helper] Errore: {e}")
        return None


def test_connessione(api_key):
    """Testa la connessione all'API provando i modelli disponibili. Ritorna (bool, messaggio)."""
    import urllib.request
    import ssl

    for model in MODELLI_PREFERITI:
        try:
            payload = json.dumps({
                "model": model,
                "max_tokens": 10,
                "messages": [{"role": "user", "content": "Rispondi OK"}],
            }).encode("utf-8")

            req = urllib.request.Request(
                "https://api.anthropic.com/v1/messages",
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                },
                method="POST",
            )

            ctx = ssl.create_default_context()
            with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            if data.get("content"):
                # Salva il modello che funziona per le prossime chiamate
                global MODEL
                MODEL = model
                return True, f"Connessione OK — Modello: {model}"
            return False, "Risposta vuota dall'API."

        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8", errors="replace")
            except Exception:
                pass

            if e.code == 401:
                return False, "API key non valida. Verifica la chiave su console.anthropic.com."
            elif e.code == 404:
                # Modello non trovato, prova il prossimo
                continue
            elif e.code == 429:
                return False, "Troppe richieste. Riprova tra qualche secondo."
            elif e.code == 400:
                # Potrebbe essere modello non supportato dalla key
                if "model" in body.lower():
                    continue
                return False, f"Errore 400: {body[:200]}"
            elif e.code == 403:
                return False, f"Accesso negato (403). Verifica i permessi della API key.\n{body[:200]}"
            else:
                return False, f"Errore HTTP {e.code}: {body[:200]}"
        except Exception as e:
            return False, f"Errore di connessione: {e}"

    return False, f"Nessun modello disponibile. Modelli provati: {', '.join(MODELLI_PREFERITI)}"


# ═══════════════════════════════════════════════════════════════
# PROMPT ENGINEERING
# ═══════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """Sei un esperto di diritto amministrativo italiano, specializzato nella redazione di atti amministrativi per enti locali.
Lavori per il Servizio Demanio Marittimo del Comune di Sanremo (Provincia di Imperia).

Il tuo compito è scrivere ESCLUSIVAMENTE la sezione "PREMESSO CHE" di una determinazione dirigenziale.

REGOLE:
- Scrivi in italiano burocratico-amministrativo formale
- Usa il linguaggio giuridico proprio degli atti della PA
- Ogni premessa deve essere un periodo compiuto che termina con punto e virgola (;)
- NON scrivere le sezioni VISTI, DATO ATTO, RITENUTO, DETERMINA — quelle sono già generate automaticamente
- Scrivi solo le premesse, una per paragrafo, separate da righe vuote
- Fai riferimento alle norme pertinenti (D.Lgs. 36/2023, D.Lgs. 267/2000, ecc.)
- Contestualizza rispetto al demanio marittimo, alle spiagge, alle opere costiere
- Sii specifico ma non inventare dati che non ti vengono forniti
- Scrivi 2-4 premesse, non di più"""


def _costruisci_contesto(tipo_determina, dati, extra):
    """Costruisce un riassunto dei dati per il prompt."""
    ctx = {
        "tipo_determina": tipo_determina,
        "oggetto": dati.get("oggetto") or dati.get("affidamento_oggetto", ""),
        "tipo_prestazione": dati.get("tipo_prestazione", ""),
        "tipo_procedura": dati.get("tipo_procedura", ""),
        "importo": dati.get("importo_affidato", dati.get("importo", 0)),
        "operatore": dati.get("ragione_sociale", dati.get("operatore_nome", "")),
        "cig": dati.get("cig", ""),
        "rup": dati.get("rup", ""),
        "rif_normativo": dati.get("rif_normativo", ""),
        "indicazioni_utente": dati.get("indicazioni_ai", ""),
    }
    # Premesse manuali esistenti (se l'utente ne ha già scritte)
    for k in ["premessa_1", "premessa_2", "premessa_3"]:
        if dati.get(k):
            ctx[k] = dati[k]

    if extra:
        ctx.update(extra)

    return ctx


def _costruisci_prompt(tipo_determina, contesto):
    """Costruisce il prompt specifico per il tipo di determina."""
    ctx_json = json.dumps(contesto, ensure_ascii=False, indent=2)
    guidance = f"\n\nISTRUZIONI SPECIFICHE DALL'UTENTE:\n{contesto.get('indicazioni_utente')}\nSEGUI SCRUPOLOSAMENTE QUESTE INDICAZIONI." if contesto.get("indicazioni_utente") else ""

    prompts = {
        "Determina a contrarre": f"""Scrivi le premesse per una DETERMINA A CONTRARRE (che include anche l'affidamento diretto).

Dati dell'affidamento:
{ctx_json}{guidance}

Le premesse devono:
1. Inquadrare le competenze del Servizio Demanio Marittimo
2. Descrivere l'esigenza che motiva l'intervento
3. Fare riferimento alla tipologia di lavori/servizi
4. Se ci sono premesse già scritte dall'utente, integrale e migliorale mantenendo il contenuto""",

        "Liquidazione": f"""Scrivi le premesse per una DETERMINA DI LIQUIDAZIONE.

Dati:
{ctx_json}{guidance}

Le premesse devono:
1. Richiamare l'affidamento originario
2. Descrivere brevemente le prestazioni eseguite
3. Attestare la regolare esecuzione""",

        "Impegno di spesa": f"""Scrivi le premesse per una DETERMINA DI IMPEGNO DI SPESA.

Dati:
{ctx_json}{guidance}

Le premesse devono:
1. Motivare la necessità della spesa
2. Inquadrare nel contesto del demanio marittimo
3. Fare riferimento alla copertura finanziaria""",

        "Affidamento": f"""Scrivi le premesse per una DETERMINA DI AFFIDAMENTO (post-gara o post-valutazione).

Dati:
{ctx_json}{guidance}

Le premesse devono:
1. Richiamare la procedura espletata
2. Descrivere l'esito della valutazione
3. Attestare la regolarità della procedura""",
    }

    default_prompt = f"""Scrivi le premesse per una determina di tipo: {tipo_determina}.

Dati:
{ctx_json}{guidance}

Scrivi 2-3 premesse pertinenti, contestualizzate al Servizio Demanio Marittimo del Comune di Sanremo."""

    return prompts.get(tipo_determina, default_prompt)


def process_document_ai(file_path, mode="determina"):
    """
    Analizza un documento (PDF/TXT) e estrae dati utili.
    mode: "determina", "affidamento", "fattura"
    """
    import os
    text = ""
    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == ".pdf":
            import fitz
            doc = fitz.open(file_path)
            for page in doc:
                text += page.get_text()
            doc.close()
        elif ext in [".txt", ".csv", ".log"]:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
        else:
            return None, "Formato file non supportato (usa PDF o TXT)"
            
        if not text or len(text.strip()) < 10:
            return None, "Documento vuoto o testo non leggibile."
            
    except Exception as e:
        return None, f"Errore lettura file: {e}"

    api_key = get_api_key()
    if not api_key:
        return {"oggetto": f"Analisi file {os.path.basename(file_path)}", "indicazioni_ai": "Configura API key per analisi intelligente."}, None

    # Prompt differenziati per modo
    if mode == "fattura":
        prompt = f"""Analizza questa FATTURA e estrai i dati contabili.
ESTRAI: Numero fattura, Data fattura, Importo totale (con IVA), Ragione Sociale Fornitore, SDI (se presente).

Documento:
{text[:8000]}

RISPONDI ESCLUSIVAMENTE IN JSON:
{{
  "numero": "...",
  "data": "AAAA-MM-GG",
  "importo": 0.0,
  "operatore_nome": "...",
  "sdi": "...",
  "note": "..."
}}
"""
    elif mode == "affidamento":
        prompt = f"""Analizza questo documento tecnico/preventivo per un AFFIDAMENTO del Comune di Sanremo.
ESTRAI: Oggetto, Operatore Economico, Importo (al netto di IVA o totale, specifica), CIG, CUP, Durata.

Documento:
{text[:8000]}

RISPONDI ESCLUSIVAMENTE IN JSON:
{{
  "oggetto": "...",
  "operatore_nome": "...",
  "importo": 0.0,
  "cig": "...",
  "cup": "...",
  "durata": "...",
  "note": "..."
}}
"""
    else: # determina
        prompt = f"""Analizza questo documento e suggerisci i dati per una determina del Comune di Sanremo.
ESTRAI: Oggetto, Operatore Economico (se presente), Importo (se presente), CIG (se presente), Numero e Data Fattura (se è una liquidazione), e scrivi 3-4 PREMESSE formali.

Documento:
{text[:8000]}

RISPONDI ESCLUSIVAMENTE IN FORMATO JSON:
{{
  "oggetto": "...",
  "operatore_nome": "...",
  "importo": 0.0,
  "cig": "...",
  "numero_fattura": "...",
  "data_fattura": "...",
  "premesse": "...",
  "note": "..."
}}
"""

    try:
        import urllib.request
        import json
        import ssl
        
        payload = json.dumps({
            "model": MODEL,
            "max_tokens": 1500,
            "messages": [{"role": "user", "content": prompt}],
            "system": "Sei un assistente amministrativo che estrae dati da documenti tecnici e contabili. Rispondi solo con JSON valido.",
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST",
        )

        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=40, context=ctx) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        if data.get("content"):
            raw_json = data["content"][0].get("text", "").strip()
            # Pulizia markdown se l'AI lo include
            if "```json" in raw_json:
                raw_json = raw_json.split("```json")[1].split("```")[0].strip()
            elif "```" in raw_json:
                raw_json = raw_json.split("```")[1].split("```")[0].strip()
            
            return json.loads(raw_json), None
            
    except Exception as e:
        return None, f"Errore AI: {e}"

    return None, "Impossibile analizzare il documento."
