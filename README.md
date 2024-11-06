# Eksamen/Privatist Script 
Løsningen leser en .xlsx eller .xls fil og flytter elever/kandidater som er elever i fylket inn og ut av en gruppe. Denne gruppen har et sett med regler som gjør at personer i gruppen vil få en rekke regler på seg og få begrenset nett tilgang. 
## Setup
Ting du trenger å sette opp for at scriptet skal fungere.
### App registration
Opprett en appregistration med følgende API permissions
1. Microsoft Graph 

|Permission Name|Type|Description|
|---|---|---|
|User.Read.All|Application|Read all users' full profiles|
|CustomSecAttributeAssignment.Read.All|Application|Read custom security attribute assignments|

### Optional 
Følgende variabler i .env er valgfritt. 
1. PAPERTRAIL_HOST
    1. Logging til papertrail 
2. PAPERTRAIL_TOKEN
    1. Logging til papertrail
3. TEAMS_WEBHOOK_URL
    1. Logging til teams, om du ønsker det.
4. NODE_ENV
    1. Denne MÅ stå til NODE_ENV=production om du ønsker logging til teams og papertrail, om du ikke legger til NODE_ENV eller om NODE_ENV står til dev vil de kun logges lokalt.


### Config
Opprett en .env fil med følgende variabler og fyll ut.

```md
    AZURE_APP_TENANT_ID=
    AZURE_APP_ID=
    AZURE_APP_SECRET=
    NETTPSERRE_EKSAMEN_GROUP_ID=GroupID of the block group
    MAIL_URL=
    MAIL_KEY=
    MAIL_SUBJECT=
    MAIL_FROM=test.test@test.no
    MAIL_TO=test.test@testfylke.no,test-test@test.com
    SERVER_PATH=Local/Nettwork path
    SCRIPT_TYPE=eksamen || privatist
    PAPERTRAIL_HOST=
    PAPERTRAIL_TOKEN=
    TEAMS_WEBHOOK_URL=
    NODE_ENV=
```

### Folder structure
Opprett en mappestruktur der du ønsker at filene skal leses fra. Denne mappen skal inneholde en "finished" folder og en "logs" folder. 

## File(s)
Scriptet leser en eller flere filer og tar også hennsyn til om det er flere "sheets" i excel filen. 

Filen/Filene kan du kalle hva som helst.

Filen/Filene må inneholde følgende felter: 
| Fødselsnummer | Eksamensdato | Eksamensparti |
|----------------|--------|--------|
| 01010112345    | 06.11.2024 | Eksamensparti1 |
| 02020223456    | 07.11.2024 | Eksamensparti2 |

## End of exam period
Når scriptet ikke finner flere fremtidige datoer i en fil vil filen bli flyttet til en "finished" mappe. 

OBS OBS! Scriptet håndterer ikke fremtidige datoer på sheet/ark nivå. Så om en fil inneholder 10 sheets/ark der 8 sheets/ark ikke har noen fremtidige datoer vil disse fortsatt bli lest. Det er derfor anbefalt å lage f.eks en fil per skole eller noe lignende. 

## Testing & running the script
Hvordan kjøre og teste scriptet. 

Hva du trenger:
1. Du har fullført alle stegene i [Setup](#setup)
2. Du har en eller flere filer med og/eller uten kandidater som er elever i din tenant på formatet i [File(s)](#files)

### Running the script 
Naviger til prosjektmappen eller gjør dette direkte. 

Send inn `prod` som et parameter (se [Testing the script](#testing-the-script)) eller ingenting. 
```PowerShell
node index.js
```
### Testing the script 
Naviger til prosjektmappen eller gjør dette direkte. 

Scriptet tar inn to argumenter `test` eller `prod`. 
```PowerShell
node index.js test
```
```PowerShell
node index.js prod
```

## Logs
I `logs` mappen du opprettet vil det genereres en logfil som gir deg en oversikt over hvem som er meldt inn, hvem som er meldt ut og hvem som ikke er elev i din tentant.

Det vil også logges error til en fil om det skulle oppstå feil. 

Om du har satt det opp vil det logges til papertrail og teams.

