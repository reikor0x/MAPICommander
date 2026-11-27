# MAPICommander

Console per Windows che usa Extended MAPI tramite interop nativa, senza EWS e senza registrare un'app Entra. Usa lo stesso profilo di **Outlook classico** (non il nuovo Outlook) e i relativi permessi Exchange. La mailbox delegata con Full Access deve apparire come **store** del profilo: aprila prima in Outlook tramite automapping o aggiungila alla configurazione del profilo. Questa versione non apre per indirizzo SMTP una mailbox non montata.

## Pubblicazione in un unico EXE

Per generare l'eseguibile da distribuire, usa l'SDK .NET 10 o successivo sul PC di sviluppo:

```powershell
dotnet publish -c Release
$app = '.\bin\Release\net10.0\win-x64\publish\MAPICommander.exe'
```

Da Visual Studio puoi usare **Pubblica > Cartella** con configurazione Release. Distribuisci solo `MAPICommander.exe` dalla cartella di pubblicazione, non dalla cartella della normale build. Il runtime .NET e le relative librerie sono inclusi: non serve installarli sul PC di destinazione. All'avvio il runtime puo' estrarre componenti nativi in una propria cartella temporanea.

L'eseguibile deve avere la stessa architettura di Outlook, non necessariamente di Windows. Il progetto pubblica per x64 per impostazione predefinita. Per Outlook a 32 bit usa `dotnet publish -c Release -r win-x86` e sostituisci `win-x64` con `win-x86` nel percorso. Outlook classico e il profilo MAPI gia' configurato restano necessari. Avvia la console nella sessione Windows del proprietario del profilo. Microsoft non fornisce un wrapper ufficiale di Extended MAPI per .NET; valida l'interoperabilita con il tuo Outlook prima dell'uso operativo.

## Esempi

Se ometti `--profile`, si apre il selettore di Outlook. Usa il nome esatto dello store mostrato da `stores`; se ci sono omonimi il programma si ferma. I nomi delle cartelle sono quelli del profilo e il percorso parte dalla radice della posta dello store.

```powershell
& $app stores --profile 'Outlook'
& $app folders --profile 'Outlook' --store 'Mailbox Delegata'
& $app folders --profile 'Outlook' --store 'Mailbox Delegata' --folder 'Posta in arrivo'
& $app scan --profile 'Outlook' --store 'Mailbox Delegata' --folder 'Posta in arrivo\Archivio' --before '2025-01-01' --online
& $app scan --profile 'Outlook' --store 'Mailbox Delegata' --folder 'Posta in arrivo\Archivio' --before '2025-01-01' --online --list > .\anteprima.tsv
```

- `scan` scrive il conteggio; `--list` scrive solo `received_utc` e `entry_id_hex` come TSV. Diagnostica e tempi sono su stderr. Gli EntryID sono identificatori opachi MAPI, non ID Internet e non vanno riutilizzati tra sessioni.
- Il filtro usa la data di ricezione (`PR_MESSAGE_DELIVERY_TIME`). Specifica `--before` (limite esclusivo) oppure `--to` (include l'intera data finale); aggiungi `--from` per un limite iniziale incluso. Le date `yyyy-MM-dd` e `d/M/yyyy` sono interpretate in UTC. Per un orario locale esplicito usa, con `--from` o `--before`, ad esempio `2025-01-01T00:00:00+01:00`. Sono inclusi `IPM.Note` e sottoclassi, esclusi bozze senza data di ricezione, riunioni, cartelle di ricerca e sottocartelle.
- `--online` chiede al provider Exchange di usare il server senza passare dalla cache locale. Senza l'opzione si segue la modalita' del profilo e la cache potrebbe non contenere tutto lo storico. Il programma non ripiega sulla cache se l'apertura online fallisce.

Per spostare in **Posta eliminata della stessa mailbox**, ripeti lo scan con `--delete` e digita `SPOSTA` dopo aver controllato il conteggio. Solo per esecuzioni deliberate e non interattive aggiungi `--yes`.

```powershell
& $app scan --profile 'Outlook' --store 'Mailbox Delegata' --folder 'Posta in arrivo\Archivio' --before '2025-01-01' --online --delete
```

Per eliminare **senza passare da Posta eliminata**, usa `--hard-delete` al posto di `--delete`. La conferma richiesta e' `ELIMINA DEFINITIVAMENTE`; `--yes` la salta. La conservazione lato server dipende dalle policy di Exchange, ad esempio retention e hold.

Con `--whatif`, anche se specifichi `--delete` o `--hard-delete`, l'app apre la cartella in sola lettura e scrive un report TSV con data di ricezione UTC ed EntryID, senza corpo, oggetto o allegati. Non combinare `--whatif` e `--yes`.

```powershell
# Anteprima dell'hard delete dal 1 gennaio 2018 al 31 dicembre 2028 inclusi (UTC)
& $app scan --profile 'Outlook' --store 'Mailbox Delegata' --folder 'Posta in arrivo\Archivio' --from '1/1/2018' --to '31/12/2028' --online --hard-delete --whatif > .\anteprima.tsv
```

Per procedere dopo aver letto il report, rilancia lo stesso comando senza `--whatif` e senza il reindirizzamento. Verra' eseguita una nuova scansione: il report non blocca la cartella. Le modifiche non sono atomiche: un errore o `Ctrl+C` puo' lasciare blocchi gia' spostati o eliminati. In caso di esito parziale, riesegui lo scan prima di riprovare.

## Prestazioni e verifica

Il filtro sulla data e' applicato da `IMAPITable::Restrict`; `QueryRows` legge fino a 512 righe alla volta, richiedendo solo data, classe e, se serve, EntryID. Nessuna apertura del singolo messaggio, lettura del corpo o download di allegati. Lo spostamento usa `CopyMessages` con `MESSAGE_MOVE`; l'hard delete usa `DeleteMessages` con `DELETE_HARD_DELETE`. Entrambi procedono a blocchi di 256, dopo aver raccolto gli ID delle corrispondenze, per non alterare la tabella durante la scansione. Tempi effettivi: misurali sulla cartella in questione; non e' garantito un vantaggio rispetto a EWS.

Per il collaudo usa una cartella di prova nella mailbox delegata: includi messaggi precedenti, esattamente alla soglia, successivi e uno in una sottocartella. Verifica conteggio e TSV, annulla una prova con `--delete` e controlla che non cambi nulla. Poi conferma lo spostamento e controlla sorgente e Posta eliminata.

Riferimenti: [accesso al profilo in MrMAPI](https://github.com/microsoft/mfcmapi/blob/main/MrMapi/MrMAPI.cpp), [tabelle di contenuto in MFCMAPI](https://github.com/microsoft/mfcmapi/blob/main/core/mapi/processor/mapiProcessor.cpp), [header di Microsoft MAPI Stub Library](https://github.com/microsoft/MAPIStubLibrary/tree/main/include).