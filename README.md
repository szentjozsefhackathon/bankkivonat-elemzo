# bankkivonat-elemzo
A project-en belül két funkciót valósítottunk meg.
## Automatizált banki letöltő modul

Az automatizált letőltő specifikus a CIB Businessonline web applikációhoz.

A letőltő program python programozási nyelven lett megvalósítva.

2 paramétere van: év, hónap

Letölti a megadott időszakra a bankszámla kivonatot, és a hozzá tartozó postai csekkes befizetésekről szóló bizonylatot PDF és UJF formátumokban. Ezeket a file-okat specifikusan a megadott év-hónap szerinti könyvtárba gyűjtjük.

## Kivonat elemző modul

A letöltött kivonat és postai csekk file-ok banki formátuma a script bemeneti paraméterei.

Formátumok: 
- Kivonat: ISO 20022 camt.053 (XML file)
- UJF: file:///home/dzsorden/Downloads/CIB%20Automata%20Lek%C3%A9rdez%C5%91%20termin%C3%A1l%20_F%C3%A1jlform%C3%A1tumok_20240403.pdf

Minden hónap külön fülre (munkalapra) kerül *ÉÉÉÉ_HH* névvel.

A név feldolgozáshoz szükség van regisztrációra itt: https://huggingface.co/NYTK/named-entity-recognition-nerkor-hubert-hungarian

Regisztráció után lehet token generálni, amit környezeti változóként kell beállítani:
- Linux: export HF_TOKEN=ide_másold_a_tokent
- Powershell: $env:HF_TOKEN="ide_másold_a_tokent"