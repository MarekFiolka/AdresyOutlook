# AdresyOutlook

Narzędzia VBA do zarządzania i normalizacji adresów w kontaktach Microsoft Outlook, ze szczególnym uwzględnieniem poprawnego mapowania pól adresowych oraz integracji z bazą kodów pocztowych.

---

## 🎯 Cel projektu

Projekt rozwiązuje problem niepoprawnego przetwarzania adresów przez Outlook (np. zamiana kolejności ulicy i numeru, błędne przypisanie województwa do miasta, itp.).

Główne założenia:

- deterministyczne przetwarzanie adresów (bez „zgadywania” Outlooka)
- poprawne mapowanie:
  - ulica
  - kod pocztowy
  - miejscowość
  - województwo
  - kraj
- automatyczne uzupełnianie województwa na podstawie kodu pocztowego
- pełna kontrola nad logiką przez kod VBA

---

## 🧱 Architektura

Projekt oparty jest o podejście obiektowe w VBA zgodne ze stylem Rubberduck:

- każda klasa używa prywatnego `Type` + zmiennej `this`
- hermetyzacja stanu
- jawna inicjalizacja i walidacja
- rozdzielenie odpowiedzialności

### Kluczowe komponenty

#### 📦 Modele danych
- `KodPocztowy` – pojedynczy rekord
- `KodyPocztowe` – kolekcja + indeks po kodzie

#### 🧠 Kontekst aplikacji
- `KontekstAplikacji`
  - lazy loading danych
  - centralny dostęp do zasobów
  - zarządzanie ścieżkami

#### 🔧 Logika
- parser CSV (UTF-8, separator `;`)
- parser adresu (heurystyki PL)
- naprawiacz adresów Outlook

---

## 📂 Struktura repozytorium

```
AdresyOutlook
│
├── cls         ' klasy VBA
├── bas         ' moduły VBA
├── docs
│   └── Kody pocztowe.csv
└── README.md
```

---

## 📮 Baza kodów pocztowych

Projekt wykorzystuje plik:

```
docs/Kody pocztowe.csv
```

Docelowa lokalizacja w środowisku użytkownika:

```
C:\Users\marek\Documents\Pliki programu Outlook\Kody pocztowe.csv
```

### Wymagania dla pliku CSV

- kodowanie: **UTF-8 BOM**
- separator: `;`
- struktura:

```
KOD POCZTOWY;ADRES;MIEJSCOWOŚĆ;WOJEWÓDZTWO;POWIAT
```

---

## ⚙️ Konfiguracja

Ścieżka do pliku CSV ustawiana jest w klasie:

```vba
KontekstAplikacji.SciezkaKodowPocztowych
```

Można ją zmienić dynamicznie:

```vba
AppContext.SciezkaKodowPocztowych = "E:\VBA\AdresyOutlook\docs\Kody pocztowe.csv"
AppContext.ReloadKodyPocztowe
```

---

## 🚀 Użycie

### Naprawa adresu aktualnego kontaktu

```vba
NaprawAdresBiznesowyBiezacegoKontaktu
```

Makro:

- analizuje adres
- normalizuje pola
- uzupełnia województwo (tylko dla Polski)
- zapisuje kontakt

---

## 🇵🇱 Logika dla Polski

Uzupełnianie województwa następuje tylko gdy:

- kraj = Polska (lub Poland)
- lub brak kraju, ale dane wskazują na adres w Polsce

Nigdy nie uzupełniamy województwa dla adresów zagranicznych.

---

## 🧪 Testowanie

Projekt jest przygotowany pod testy Rubberduck:

- parser CSV
- parser adresu
- baza kodów pocztowych
- naprawiacz adresów

---

## 🔧 Wymagania

- Microsoft Outlook (VBA)
- Windows (testowane w środowisku PL)
- Rubberduck VBA (zalecane)

---

## 📌 Uwagi

- Outlook nie posiada poprawnego parsera adresów dla PL — projekt go zastępuje
- plik CSV nie powinien być otwarty w Excelu podczas działania (blokada pliku)
- dane są ładowane leniwie (lazy loading)

---

## 📈 Kierunki rozwoju

- walidacja kod ↔ miejscowość
- tryb preview (bez zapisu)
- obsługa wielu rekordów dla jednego kodu
- GUI wyboru adresu
- pełna integracja z Rubberduck test framework

---

## 👤 Autor

Projekt rozwijany w celach praktycznych – zarządzanie kontaktami i automatyzacja Outlook.

---

## 📄 Licencja