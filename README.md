# In bulk omzetten van RTF naar DOCX

In 2018 gaf Centric aan dat het om verschillende redenen verstandig is om van RTF over te stappen naar OOXML (dus DOCX). Nu kan je een RTF bestand openen, opslaan als DOCX en weer sluiten. Dit is niet efficiënt en geeft een risico dat je zo lang bezig bent dat het tijdens productietijd moet. In dat geval loop je kans dat de body al wel omgezet is naar DOCX, maar het sjabloon nog niet. Dit kan gemakkelijker.  

In dit voorbeeld gaan we onderstaande lijst omzetten:

<img src='./img/Om_te_zetten_bestanden.png' alt = 'Een windows mappenstructuur met om te zetten rtf bestanden'>

## Installatie

Er is niet direct een installatie nodig. Het belangrijkste is het templatebestand *Converteer van RTF naar DOCX.docm*. Hiermee kan je aan de slag. Mocht je de code willen aanpassen, vind je in dezelfde map ook de broncode terug (broncode.bas).

## Handleiding

1. In de map template van deze repository staat het bestand *Converteer van RTF naar DOCX.docm*. Open dit bestand.

2. Dubbelklik op de tekst *DUBBELKLIK HIER*. Zorg er wel voor dat macro's ingeschakeld zijn door op *inhoud inschakelen* te klikken in de waarschuwingsbalk als je die krijgt.

    <img src='./img/dubbelklik_hier.png' alt = 'Een eerste blik op de template'>

3. Ga naar het pad waar de RTF documenten staan die omgezet moeten worden en klik op *OK*

    <img src='./img/kies_map.png' alt = 'Een venster om een map te kiezen'>

4. Wacht rustig af. Onder in het scherm zie je bij welk bestand het systeem is. Als dat scherm verdwijnt is het systeem klaar (er staat nu weer Pagina 1 van 1 en 2 van 2 woorden)

    <img src='./img/conversie_bezig.png' alt = 'Conversie bezig'>

In stap 3 heb je een map gekozen waarin de RTF bestanden staan. In diezelfde map zijn nu de DOCX bestanden geplaatst:

<img src='./img/omzetting_voltooid.png' alt = 'Conversie klaar'>