# MsWordVersioning
MS Word script for (basic) version control of documents.

## Så fungerar versionshanteringen

Det finns ett makro för att skapa nya undervserioner av ett dokument (exempelvis 1.3 till 1.4) och ett för att skapa ny huvudversion (exempelvis 1.4 till 2.0). När något av dessa makron körs händer följande:

* Du ombeds ge en kort beskrivning av ändringarna i dokumentet, exempelvis "Justeringar efter diskussioner 2017-10-25" eller "Version som beslutats av enhetschef".
* Versionshistoriken i dokumentet får en ny rad, med den beskrivning du gett.
* En pdf och ett låst Word-dokument sparas i en underkatalog, för att bevara hur dokumentet såg ut just i denna version. Word-dokumentet kan kommenteras, men inte redigeras på annat vis.

För att kunna använda makrona behöver du spara ditt dokument som makro-aktiverat dokument. Detta gör du första gången du sparar dokumentet, i valet av filformat (under filnamnet).

## Några detaljer och tips

* För att komma åt makron på ett smidigt sätt rekommenderas att lägga till knappar i menyraden. Det gör du genom att högerklicka i menyraden och välja "anpassa menyfliken". I rutan som dyker upp kan du filtrera fram makron i vänsterspalten, och lägga in dem på lämpligt ställe i menystrukturen i högerspalten. Det går också att byta namn och ikoner för knapparna.
* För att faktiskt kunna köra makron krävs dels att du sparat dokumentet som ett makro-aktiverat Word-dokument, dels att säkerhetsinställningarna tillåter att du kör makron. Säkerhetsinställningarna kan du ändra under arkiv -> alternativ -> säkerhetscenter -> inställningar för säkerhetscenter.
* Versionshistoriken kan flyttas till sist i dokumentet, och redigeras fritt, utan att makrot påverkas. Det ligger dock ett osynligt bokmärke vid historiken, och om hela historiken raderas försvinner även bokmärket och makrot kommer inte längre kunna lägga till nya rader.
* Som standard läggs exporterade dokument i en mapp med samma namn som dokumentet. Det finns ytterligare ett makro ("bytExportkatalog"), som kan användas för att ändra namnet på mappen som ska användas.
