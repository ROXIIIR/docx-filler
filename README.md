Acest proiect implementează un Agentic AI care completează automat un fișier DOCX multi-pagină folosind date dintr-un fișier JSON. Soluția respectă integral layout-ul, structura și formatarea documentului original.

Implementarea este realizată exclusiv în Python și folosește:

python-docx pentru manipularea fișierelor DOCX

OpenAI GPT-4.1-mini pentru mapping agentic între etichetele din DOCX și valorile JSON

logică avansată pentru detectarea și completarea câmpurilor complexe

Funcționalități principale
1. Detectarea câmpurilor de completat

Agentul identifică automat zonele din document care trebuie completate, inclusiv:

blank-uri simple (_____, .....)

multiple blank-uri în același paragraf

blank-uri în structuri complexe precum:

______ zile, respectiv până la data de ________
(durata în litere și cifre)      (ziua/luna/anul)


câmpuri fără etichetă directă (prin mecanismul last_label)

contexte textuale (de exemplu: „până la data de”)

2. Reasoning agentic cu LLM

LLM-ul primește:

lista etichetelor extrase din DOCX

datele brute din JSON

Modelul produce un mapping complet între fiecare etichetă și valoarea din JSON care se potrivește semantic. Dacă nu există corespondență, întoarce un șir gol.

3. Completarea automată a documentului

Documentul este completat fără a schimba formatarea:

se păstrează fonturile

se păstrează stilurile și alinierea

se păstrează structura tabelelor

paragrafele se reconstruiesc printr-un algoritm dedicat care păstrează structura Word

Structura proiectului
n8n-docx-filler/
│
├── agent_docx_filler.py          # Agentul AI complet
├── fill_docx_from_json.py        # Versiunea rule-based simplificată
├── sample_forms.docx             # Documentul DOCX original
├── input_date.json               # Datele pentru completare
├── sample_forms.llm_filled.docx  # Output completat de agent
├── README.md                     # Acest document
└── Exercise - Fill a Multi-Page DOCX.pdf

Cum se rulează
Instalare dependențe
pip install python-docx openai

Configurare cheie API
setx OPENAI_API_KEY "CHEIA_TA"

Rulare agent
python agent_docx_filler.py --template sample_forms.docx --data input_date.json --output sample_forms.llm_filled.docx

Îmbunătățiri aduse

Pentru a face agentul compatibil cu documente complexe, au fost implementate următoarele:

1. Extinderea logicii de parsare

detectarea și procesarea mai multor blank-uri succesive în același paragraf

detecție pentru blank-uri cu paranteze

generarea de etichete contextuale

suport pentru câmpuri fără etichetă (prin last_label)

2. Reconstrucție inteligentă a paragrafelor

Funcția de completare înlocuiește valori în blank-uri reconstruind paragraful fără a afecta stilurile Word.

3. Extinderea JSON-ului

Au fost adăugate câmpuri suplimentare pentru:

valabilitatea ofertei (zile numeric, zile în litere)

data finală

date calendaristice necesare câmpurilor cu paranteze

reprezentanți, funcții, autorități, detalii contractuale

Aceste date permit LLM-ului să completeze inclusiv cele mai complexe câmpuri.