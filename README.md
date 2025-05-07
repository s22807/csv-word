# Manual
## Przenoszenie instalacji
W celu przeniesienia instalacji (katalogu tej aplikacji) należy skorzystać z pliku requirements.txt okreslajacego zaleznosci skryptu
Można skorzystać z python.exe znajdujacego sie w katalogu ./.venv (nie trzeba pobierac z internetu)

Kroki

Kroki:
1. Pobierz interpreter python ze strony https://www.python.org/downloads/ i zainstaluj. Ważne żeby na początku instalacji dodać python.exe do ścieżki (PATH). Wtedy nie trzeba wpisywać dokładnej lokalizacji programu, jedynie nazwę python.
2. Utwórz katalog docelowy projektu, lokalizację do której będziesz chciał pobrać projekt. 
3. Przejdź do katalogu docelowego uruchamiajac w nim terminal cmd lub uruchamiając w dowolnej lokalizacji i wykonując polecenie `cd <katalog projektu>`
4. Następnie uruchom polecenie `git clone https://github.com/s22807/csv-word.git`
5. W terminalu wykonaj (znajdujac sie w katalogu projektu): `python -m venv <katalog docelowy>`
6. `<katalog docelowy>\Scripts\activate`
7. `pip install -r .\requirements.txt`
Gotowe. Python z zaleznosciami bedzie mogl byc uruchamiany z nowej lokalizacji
## Uruchamianie
Główny skrypt programu main.py należy uruchomić za pomocą wirtualnego srodowiska Python, żeby skorzystać z dociągnietych zależnosci (biblioteki pandas i python-docx)
Kroki:
1. W terminalu wykonaj (znajdujac sie w katalogu projektu): `.\.venv\Scripts\activate`
2. python .\main.py
Dokumenty wynikowe pojawią się w głównym katalogu projektu

## Konfiguracja
Skrypt można skonfigurować za pomocą pliku config.py. Jego struktura jest listą key=value. Konfigurowalne parametry:
1. plik_csv (wymagany) - wskazuje na plik danych w formacie .csv, którego separatorem jest srednik (;)
2. word_template (wymagany) - wskazuje na spreparowany dokument Word do którego będzie wstawiany tekst z pliku danych
3. dirs (opcjonalny) - jesli chcemy zeby pliki generowaly sie w okreslonym katalogu (relatywnym do projektu lub parametru absolute_dir jesli okreslony) to musimy uzupelnic ten parametr. Parametr pobiera listę dzięki czemu możemy zdefiniować więcej niż jedno rozgałęzienie katalogów. Jeżeli w liscie pojawi się nazwa wystepujaca w nagłówkach pliku danych (.csv) to nazwa katalogu bedzie utworzona na podstawie wartosci pola tej kolumny dla generowanego pliku Word. Brak parametru poskutkuje utworzeniem plikow Word w katalogu glownym projektu.
4. absolute_dir (opcjonalny) - jesli chcemy zmienic lokalizacje tworzonych katalogow i plikow Word tak, zeby nie tworzyly sie w katalogu glownym projektu to mozemy tu wpisac pełna sciezke. Brak parametru poskutkuje utworzeniem katalogów i/lub plikow Word w katalogu glownym projektu. 

## Przygotowanie wzoru worda
Plik Word na podstawie ktorego maja byc generowane dokumenty do druku musza byc spreparowane w taki sposob aby skrypt mogl powiazac z nim pola z pliku z danymi. Do pustego wzoru w miejsca w których maja sie pojawić dane z pliku z danymi należy w nawiasach klamrowych umiescic nazwe odpowiadajacego mu pola z pliku z danymi. W pliku Word mozna wielokrotnie umiescic ta sama nazwe pola, kazde bedzie odpowiednio podmienione.
Szablon worda musi byc wskazany w pliku config.py

## Przygotowanie danych do podmiany
Plik danych musi byc z excela wyeksportowany do pliku .csv z separatorem srednikiem (;) i wskazany w pliku config.py