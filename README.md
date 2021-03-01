# Manuale utente del software LEX - Lettore ed Estrattore_xls-xlsx

Il software consente all’utente di leggere dei file con estensione .xlsx o .xls ed estrarre da essi dei dati specifici che verranno proposti in output come un file di tipo .csv.

E’ importante che prima di utilizzare il software si personalizzi il codice in vista dei propri interessi. Nello specifico:
Rinominare i propri file con la nomenclatura seguente: un file riferito al gennaio del 2015 dovrà essere rinominato 1_2015 e così tutti gli altri.

Tutti i file dovranno essere inseriti in un'unica cartella chiamata DATI.

Personalizzare la lista_anni e la lista_mesi in base al proprio interesse: 
`lista_anni = np.arange(2015,2021)` - Inserire nella parentesi il periodo di tempo che coprono i file per quanto riguarda l’anno assicurandosi che il secondo numero sia avanti di un anno: in questo caso se i miei file vanno dal 2015 al 2020 allora inserisco 2015 e 2021.

`lista_mesi = np.arange(1,13)` - Inserire nella parentesi il periodo di tempo che coprono i file per quanto riguarda il mese assicurandosi che il secondo numero sia avanti di un mese: in questo caso se i miei file vanno da gennaio (1) a dicembre (12)  allora inserisco 1 e 13.

Personalizzare la funzione di configurazione in base agli anni e i mesi dei propri file e scegliere i parametri che si vogliono ricercare. 

Quelli utilizzati di default nel software sono:
* `ext` , che indica l’estensione del file 
* `tab` , che indica la tabella del documento che si vuole utilizzare (se il documento presenta più tabelle).
* `usecols` , ovvero le colonne che si desidera utilizzare. Per esempio A,C:F,K:N,S:V indica che si vogliono utilizzare la colonna A, la colonna C, le colonne da F a K e le colonne da S a V.
* `skiprows` , che indica quante righe del documento il software deve saltare prima dell’inizio della lettura
* `names` che indica il modo in cui rinominare le colonne nell'ordine in cui sono nel documento di input

Ecco il link per visualizzare tutti i parametri disponibili della funzione [pandas read_excel](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html ) 

Per informazioni più dettagliate sulla funzione read_excel ecco un [tutorial](https://www.youtube.com/watch?v=xofx-WqzYrU)

Inserire il nome della colonna di interesse nella variabile `colonna`

Inserire il nome della riga di interesse nella variabile `valore`

Nel caso in cui nel susseguirsi dei documenti la riga interessata abbia di più di un nome si può utilizzare e personalizzare la variabile `valore2`. Se invece il nome è sempre uno inserire nella variabile `valore2` lo stesso nome della variabile `valore`.

Inserire il percorso dei propri file nella riga nel quale viene specificato:  `file = "/"`

Nella riga `df = pd.read_excel(xl,....)` personalizzare i parametri in base a quelli utilizzati nella funzione di configurazione facendo attenzione a lasciare nella parentesi `xl`.

Nell’ultima riga inserire il nome che si desidera per il proprio documento csv

