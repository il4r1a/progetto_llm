# IL GENERE SECONDO UN LLM

1. DOMANDA DI RICERCA

I modelli linguistici di grandi dimensioni apprendono le proprie associazioni statistiche da enormi quantità di testo prodotto da esseri umani. Questo testo contiene, inevitabilmente, gli stereotipi sociali presenti nella cultura in cui è stato scritto — inclusi quelli di genere legati alle professioni. Quando un modello viene interrogato su una persona descritta solo tramite la propria attività lavorativa, tende quindi a produrre un’associazione di genere anche in assenza di qualunque informazione reale sul genere di quella persona.
Il progetto nasce per misurare in modo controllato e riproducibile questo fenomeno: non per dimostrare che il bias esiste (è già ampiamente documentato in letteratura), ma per capire quali dimensioni della descrizione (competenza tecnica, cura/relazione, autorità/leadership) lo rafforzano, lo attenuano o lo lasciano invariato, e quanto questo comportamento sia stabile.
La domanda di ricerca è dunque così formulata: In che modo un modello linguistico associa un genere a una persona descritta professionalmente, e come cambia questa associazione quando alla professione si aggiungono informazioni su competenza tecnica, componente relazionale/di cura o ruolo di leadership?

Le ipotesi di partenza, testate nel progetto, sono:
a. Le professioni fortemente stereotipate (es. carpenter, electrician, engineer) producono attribuzioni di genere quasi totalmente uniformi. 
b. L’aggiunta di dettagli tecnici o relativi posizioni manageriali aumenta la quota di attribuzioni maschili. 
c. L’aggiunta di dettagli relativi a cura, relazione, supporto aumenta la quota di attribuzioni femminili. 
d. L'alterazione della temperatura di campionamento non dovrebbe capovolgere la direzione media dell’associazione, se lo stereotipo è “reale” e non rumore statistico.

2. METODOLOGIA

- uso di llm in locale (ollama), con sessioni indipendenti
- 15 professioni, scelte per coprire tre fasce: stereotipicamente maschili, stereotipicamente femminili e ambigue/dipendenti dal contesto.
- 4 condizioni informative: professione, professione + compito tecnico, professione + compito relazionale, professione + leadership
- esecuzione dell'esperimento con due temperature: 0 e 0,8
- scelta forzata tra A/B, con corrispondenza a "male" e "female" randomizzata
- 4 condizioni di controllo
- 20 ripetizioni per ogni condizione; 10 ripetizioni per le condizioni di controllo

3. OSSERVAZIONI

Il modello mostra un bias di genere occupazionale di base consistente, che, generalmente:
- si rafforza quando la descrizione enfatizza competenza tecnica;
- si attenua quando la descrizione enfatizza cura/relazione;
- non si sposta in modo uniforme con la leadership, perché il suo effetto dipende fortemente dalla professione considerata;
- è per alcune professioni “saturo”, cioè talmente forte da risultare insensibile a qualsiasi attributo aggiunto;
- è sostanzialmente stabile rispetto alla temperatura di campionamento, quindi non riducibile a rumore statistico;
- è presente come bias di default verso il maschile anche in totale assenza di informazioni occupazionali (condizioni di controllo)
