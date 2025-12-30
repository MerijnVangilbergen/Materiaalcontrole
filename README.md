# Materiaalcontrole
Dit is een simpele GUI om snel aan het begin van de les:
- te overlopen en in te geven welke leerlingen in orde zijn,
- te visualiseren hoe vaak leerlingen niet in orde waren, (Elke leerling heeft 3 'levens'. Toon dit gerust op groot scherm zodat de leerlingen het zien.)
- een sanctie toe te dienen aan leerlingen die driemaal niet in orde waren.

Optioneel:
- Een dubbele overtreding (rood i.p.v. oranje) kan ingegeven worden als een leerling betrapt wordt op een leugen.
- Een vakje kan aangevinkt worden om te noteren dat de leerling een nota heeft gekregen, die les erop gecontroleerd moet worden.
- Aan het einde van een trimester kan je sancties archiveren en alle leerlingen weer 3 'levens' geven.
- Je kan manuële controle of correcties doorvoeren in Materiaalcontrole.xlsx. Alle data wordt hierin opgeslagen. (Zorg er altijd voor dat de excel gesloten is als je het programma start. Je zal op een error lopen als dit niet het geval is.)

# Manual Installation
![alt text](<HowToDownload.png>)
- On github, press Code > download zip.
- Unzip this folder.
- Open run.bat a first time. It will install python locally.
- Open run.bat a second time. It will create a virtual environment for the programme and download the required python packages. Once finished, it opens the GUI.
- When you open run.bat a third time, it will activate the previously created environment and open the GUI quickly.
- (Optional) You can't move run.bat to the desktop for quick access, but you can create a shortcut (snelkoppeling) and move that to the desktop.

# Gebruiksaanwijzing
Stappenplan:
- Installeer zoals hierboven beschreven.
- Open het bestand Materiaalcontrole.xlsx.
- Creëer een sheet voor elke klas.
- Geef de namen van alle leerlingen in.
- Zorg ervoor dat alle kolommen zijn zoals in de voorbeeld sheets. De kolom 'Nota' is optioneel.

Open het programma door run.bat te openen. \
Open je run_archive_sanctions.bat, dan:
- wordt er een dated copy Materiaalcontrole_yyyymmdd.xlsx gemaakt van de excel,
- wordt de oorspronkelijke Materiaalcontrole.xlsx geupdatet door alle leerlingen weer 3 'levens' te geven.