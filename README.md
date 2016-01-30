# tm2ppt
Zet Toernooimanager-output om in een Powerpoint-presentatie

## Inleiding

Dit tooltje zet eenvoudig uitvoer van Toernooimanager om in een Powerpoint-presentatie.
Aan de hand van .csv-uitvoer van Toernooimanager wordt door een Python-bestand een template-presentatie automatisch gevuld met de standen van een (Zwitsers) toernooi.
Dit is met name geschikt voor schooldamtoernooien, omdat er daar niet altijd tijd is de uitslagen in (bijvoorbeeld) Toernooibase in te voeren.

## Werking

Het bestand `process.py` leest .csv-uitvoer van Toernooimanager in (voorbeeld: `ronde1.CSV`) en zet dit in een template-.pptx-bestand (voorbeeld: `Standen.pptx`).

### Toernooimanager
Voer de standen uit in .csv-formaat door in Toernooimanager te kiezen voor Uitvoer => Alle ranglijsten en daar te kiezen voor "Stand" bij "Serie ranglijsten". 
Vink bij "Afdrukken van" alles uit en kies voor "Disk" en dan bij "Opslaan als" voor "Kommagescheiden bestanden".

### Powerpoint-template

### Python
