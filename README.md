# Handleiding voor de leraar — Les 05: Alles op een rij

**Vak:** Toegepaste Informatica
**Doelgroep:** de ORLO-klassen — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd (formatief)
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel (vervolg op les 02–04)
**Lokaal:** 18 — computers met Windows 11, Google Workspace in Chrome
**Kernleerplandoelen:** `BK2_02.05` / `BK2_02.05.02` (tekstverwerking, minimale inhoud *tabellen*) en `BV2_04.02` (digitale inhouden creëren) — toepassen
**Deadline:** vrijdag 2 oktober 2026, 20.00 uur

De leerling zet de leveringen van een dag uit een briefje in een tabel, past de tabel aan
(rij weg, rij erbij, kolom erbij), maakt ze op, en maakt daarna zelfstandig een deelnemerslijst
voor de infosessie van de jobstudenten uit les 04. Hij levert alleen het werkdocument in.

---

## 1. Inhoud van het pakket

```text
W05 - Les 05 - ORLO - Tekstverwerking - tabellen/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 8 klassikale dia's voor de lesstart en "Ik doe"
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/
│   ├── novadepot-logo.svg/.png, novadepot-icon.svg   # logo NovaDepot (eigen werk)
│   ├── dalton-gent-logo.png
│   ├── fonts/                   # Atkinson Hyperlegible + Montserrat (OFL, zelf gehost)
│   └── screenshots/             # knop-uitlijnen.png (jouw schermafbeelding uit les 03) + LEESMIJ.md
├── werkdocument/
│   ├── TV4_Tabellen.docx        # het werkdocument dat de leerling INLEVERT
│   └── maak_werkdocumenten.py   # maakt het werkdocument opnieuw (python-docx)
├── lesvoorbereiding.md          # volgens §9.2 van de AI-lesplanner v2.1
├── dalton-lesfiche.html         # Dalton-lesfiche in de kleurcode: openen, Kopieer, plakken in je planner
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Geen bronbestanden en dus geen zip-bestand: het briefje en het mailtje staan in het werkdocument.

### 1b. Hoe de lespagina werkt

Dezelfde opbouw als les 04 (AI-lesplanner v2.1): links de **route** (zes stappen in twee groepen:
*Deel A: de leveringen* · *Deel B en inleveren*), midden **één stap** met vijf vaste blokken,
rechts de **checklist** met 21 concrete taken. Op een smal venster staat alles onder elkaar en zie
je alleen de taken van de huidige stap; in stap 6 staat de hele lijst open. Hoe de leerlingen hun
vensters schikken, kiezen ze zelf: de pagina zegt er niets over.

- De **theoriekaart** heeft tien kaartjes, met een kleine voorbeeldtabel waarin de koprij, een rij,
  een kolom en een cel elk een eigen kleur én een uitleg in woorden krijgen.
- **Stap 5** toont de **tabelkaart** van NovaDepot (zoals de huisstijlkaart in les 03): de leerling
  maakt de deelnemerslijst met alleen die kaart.
- **Alleen Windows 11**: alle klassen werken in lokaal 18. Alles gebeurt in de browser; de pagina
  vraagt niet naar het toestel.
- **Classroom is het vertrekpunt**: de leerlingen starten elke les in Classroom en vinden daar de
  opdracht met de lespagina en hun werkdocument. De pagina zegt alleen "in de opdracht" waar ze er
  iets uit nodig hebben (*Inleveren*).

`localStorage` bewaart alleen de vinkjes en de laatste stap (voorvoegsel `novadepot_tv4_v1_`).

---

## 2. Klaarzetten (± 15 minuten)

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 27-09-2026*

Op jouw vraag gepubliceerd, vóór je inhoudelijke controle, zodat je het op je gsm kan nakijken.
Pages staat op branch `main`, map `/ (root)`:

- **Repository:** <https://github.com/jonasdaltongent/Tabellen-NovaDepot>
- **Lespagina voor de leerlingen:** <https://jonasdaltongent.github.io/Tabellen-NovaDepot/>
- **Dia's voor het bord:** <https://jonasdaltongent.github.io/Tabellen-NovaDepot/presentatie.html>

Dat adres staat al op dia 7 en in `lesdoelen.json` (veld `bron`). Deel met de leerlingen altijd
het **Pages-adres**, niet de repositorylink. Wil je na je controle iets veranderen, zeg het: ik pas
het aan en push opnieuw. Een wijziging staat 1 à 2 minuten na de push online.

### Stap 2 — Het werkdocument omzetten en nakijken

Classroom maakt de kopie per leerling alleen in Google-formaat
([Classroom-help](https://support.google.com/edu/classroom/answer/6020265?hl=nl)), dus het
werkdocument moet een Google-document zijn.

**Eenmalig** (deed je het al voor les 04, dan hoeft het niet opnieuw): zet in Google Drive de
instelling aan die een Word-bestand bij het uploaden meteen omzet naar Google Documenten. Ga naar
[drive.google.com/drive/settings](https://drive.google.com/drive/settings) en vink **Uploads
converteren naar de indeling van een Editor van Google Documenten** aan
([Drive-help](https://support.google.com/drive/answer/2424368?hl=nl)). Let op: vanaf dan wordt élk
Word-, Excel- of PowerPoint-bestand dat jij uploadt een Google-bestand.

Daarna, voor deze les:

1. Upload `werkdocument/TV4_Tabellen.docx` **in Drive zelf**: **Nieuw** › **Bestanden uploaden**.
   Staat er geen `.docx` meer achter de naam? Dan is het een Google-document.
2. **Kijk het omgezette document na:**
   - [ ] de tabel in deel A heeft **vier even brede kolommen** en randen; *Koffiebranderij De
     Bonenbaas* past niet op één regel (dat is de bedoeling: stap 4);
   - [ ] de koprij *Uur · Leverancier · Pallets · Poort* is **niet** vet en niet gekleurd;
   - [ ] rij 1 is ingevuld (*7.30 · Bakkerij Korstjes · 2 · 3*), rij 2 tot 6 zijn leeg;
   - [ ] het briefje, de wijzigingen en het mailtje hebben een lichtgele achtergrond (mag ook
     wegvallen; de tekst is wat telt);
   - [ ] bovenaan elke pagina de koptekst, onderaan *Pagina 1*, *Pagina 2* …

> [!IMPORTANT]
> Voeg het werkdocument **niet** toe met **Uploaden** in Classroom. Die knop volgt de Drive-instelling
> niet: het bestand blijft dan een `.docx` (getest op 27-09-2026).

### Stap 3 — Eén opdracht in Google Classroom

Voeg het werkdocument toe met **Bijvoegen** › **Drive** en kies **Een kopie maken voor elke
leerling**: elke leerling krijgt een eigen kopie met de eigen naam in de titel. Voeg de lespagina
toe met **Link**.

| | Opdracht: **Tekstverwerking 4 — Alles op een rij** |
|---|---|
| **Onderwerp** | Tekstverwerking – de basis |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `TV4_Tabellen` (Google-document) — **Een kopie maken voor elke leerling** |
| **Punten** | zonder cijfer (formatief) |
| **Deadline** | vrijdag 2 oktober 2026, 20.00 uur |

> [!NOTE]
> *Een kopie maken voor elke leerling* kan je alleen kiezen **vóór** je de opdracht post.

Instructietekst (kopieer):

```text
1. Open de lespagina (link) en je werkdocument TV4_Tabellen.
2. Volg de stappen op de lespagina. Je maakt twee tabellen: deel A en deel B.
3. Klaar? Klik op Inleveren.
```

Wie niet klaar is, levert toch in: dat zeg je mondeling op het einde van de les (notities bij dia 8).
Het staat bewust niet op de lespagina en niet in de instructietekst.

Je e-mailadres is deze les niet nodig.

### Afvinklijst vóór de les

- [x] De lespagina is gepubliceerd en het adres op dia 7 klopt. (Getest op 27-09-2026.)
- [ ] Het werkdocument is een **Google-document** (geen `.docx` achter de naam), toegevoegd met
  **Drive**, en ziet eruit zoals in stap 2 hierboven.
- [ ] Een testleerling krijgt een eigen kopie met de eigen naam in de titel.
- [ ] Met die testleerling: rechtsklik in een cel toont **Rij onder invoegen** en **Rij verwijderen**;
  met de koprij geselecteerd staat **Achtergrondkleur** in de werkbalk (of achter ⋮).
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities.

### 2b. Nagelezen klikpaden (26-09-2026)

Uit de Nederlandse helppagina's, ruwe tekst. Waar jouw scherm anders zegt, geldt jouw scherm.

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Tabel invoegen | **Invoegen** › **Tabel**, kies hoeveel rijen en kolommen (max. 20 × 20) | [Docs 1696711](https://support.google.com/docs/answer/1696711?hl=nl) |
| Rij of kolom toevoegen | rechtsklik op een cel › **Kolom links invoegen** · **Kolom rechts invoegen** · **Rij boven invoegen** · **Rij onder invoegen** | idem |
| Verwijderen | rechtsklik › **Kolom verwijderen** · **Rij verwijderen** · **Tabel verwijderen** | idem |
| Kolombreedte | muisaanwijzer op de rasterlijn tot een dubbele pijl, dan slepen | idem |
| Celkleur | cellen selecteren › in de werkbalk **Achtergrondkleur** | idem |
| Tabelopties | **Opmaak** › **Tabel** › **Tabelopties**, of rechtsklik › **Tabelopties**; onder *Rij* een hoogte, **OK** | idem |
| Koprij vastzetten | rechtsklik › **Kop vastzetten tot deze rij** (vastgezette rijen sorteren niet mee) | idem |
| Sorteren | rechtsklik › **Tabel sorteren** › **Tabel sorteren in oplopende volgorde** | idem |
| Vet · centreren | **Ctrl + B** · **Ctrl + Shift + E**; werkbalk **Uitlijnen** | [Sneltoetsen](https://support.google.com/docs/answer/179738?hl=nl) · [Docs 1663349](https://support.google.com/docs/answer/1663349?hl=nl) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** (**Privéreacties** › **Posten** alleen voor jouw feedback, niet op de lespagina) | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |
| Uit les 04 | **Invoegen** › **Pagina-elementen** › **Koptekst** · **Ctrl + Enter** · **Bestand** › **Downloaden** | zie README van les 04, §2b |

**Beschreven in plaats van benoemd** (niet in de helppagina's): hoe het raster van *Invoegen ›
Tabel* eruitziet (*"beweeg over de vakjes: 3 naar rechts en 6 naar beneden"*), de namen van de
kleuren bij *Achtergrondkleur* (*"een lichtblauwe kleur"*) en het veld voor de rijhoogte in
*Tabelopties*.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1–2 lesstart (3'): retrieval van les 04 (pagina-einde, koptekst). Dia 3–6 demo (5'): lesdoel, van briefje naar tabel, en twee dingen voordoen: rij 1 invullen + het rechtsklikmenu, en de koprij selecteren en opmaken. Dia 7 *Zo werk je verder* (2') |
| **Keuzewerktijd** | **40'** | Dia 7 blijft staan; de leerlingen werken stap 1 tot 6 af, inleveren inbegrepen. Dia 8 in de laatste minuut: waarom een tabel en geen spaties? |

Keuzewerktijd = 50 minuten − instructietijd. De minuten per stap staan op dia 7 en in
`dalton-lesfiche.html`. Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert.

**Eerste rondgang, kijk naar twee dingen:**

1. Werkt iedereen in de **eigen kopie** van het werkdocument?
2. Staat er in elke cel maar **één ding**? Wie hele zinnen in de tabel typt, loopt vast in stap 3
   en 4.

**Tweede rondgang, rond stap 3:** is de juiste rij verwijderd (Frisdranken Bubbel, 9.00 uur), en
staat 11.15 uur tussen 10.30 en 13.00? Een verkeerde rij verwijderd: *Ctrl + Z*.

---

## 4. Verbetersleutel

### Deel A na stap 3 en 4

| Uur | Leverancier | Pallets | Poort | Afgetekend |
|---|---|---|---|---|
| 7.30 | Bakkerij Korstjes | 2 | 3 | |
| 8.15 | Papierhandel Vellekens | 4 | 4 | |
| 10.30 | Koffiebranderij De Bonenbaas | 1 | 4 | |
| 11.15 | Tuincentrum Groenvinger | 2 | 3 | |
| 13.00 | Speelgoed Tolletje | 3 | 3 | |
| 14.45 | Schoonmaak Glanzmann | 2 | 4 | |

- Zes leveringen, op volgorde van het uur; **geen** Frisdranken Bubbel (9.00 uur); geen lege rijen.
- Koprij vet en lichtblauw (over alle vijf de cellen); *Pallets* en *Poort* gecentreerd; elke naam
  op één regel; *Afgetekend* leeg.
- Een tijd als `8u15` of `08.15` is goed. Woorden als *om*, *uur* of *pallets* in de cellen: niet.

### Deel B

- Een titel boven de tabel, vet: *Deelnemerslijst infosessie jobstudenten*.
- Een tabel van 3 kolommen × 6 rijen: *Naam · Afdeling · Handtekening*, en de vijf jobstudenten
  (Lina Haddad – magazijn · Emre Yilmaz – magazijn · Noor Verbeke – onthaal · Kobe De Smet –
  verzending · Ayla Janssens – onthaal). De volgorde uit het mailtje of alfabetisch: allebei goed.
- Opgemaakt zoals de tabelkaart: koprij vet en lichtblauw, tekst links, geen lege rijen, elke naam
  op één regel. De kolom *Handtekening* blijft leeg.

### Vragen

| Vraag | Waar het om gaat |
|---|---|
| 1 | **7 pallets** (4 + 1 + 2). Sneller in de tabel: je kijkt alleen naar de kolommen *Pallets* en *Poort*, in plaats van zes zinnen te lezen. |
| 2 | Zodat je meteen ziet wat er in elke kolom staat; de koprij valt op en is geen gegeven. |
| 3 | Spaties en tabs schuiven op zodra een naam langer wordt of het lettertype verandert; een tabel houdt elke kolom op zijn plaats. |
| 4 | Vrij. |

**Vraag 1, het getal.** Naar poort 4 gaan Papierhandel Vellekens (4), Koffiebranderij De Bonenbaas
(1) en Schoonmaak Glanzmann (2): **7 pallets**. Vóór en na stap 3 is dat hetzelfde, want de
wijzigingen gaan alleen over poort 3. Het gaat vooral om *waar* je het sneller vindt. Tellen ze
leveringen in plaats van pallets (3), vraag dan door: *"Hoeveel pallets, niet hoeveel
vrachtwagens?"*

### Essentiële fouten — geef hier altijd feedback op

- Hele zinnen of meerdere gegevens in één cel.
- De verkeerde levering verwijderd, of de levering van 11.15 uur achteraan gezet.
- Kolommen gemaakt met spaties of tabs in plaats van een tabel (deel B).
- Lege rijen laten staan.
- Deel B zonder tabel, of de titel in de tabel in plaats van erboven.

**Feedback:** één top en één tip als privéreactie in Classroom. Verbeteren en opnieuw inleveren
mag (*Inleveren ongedaan maken*).

---

## 5. Schermafbeeldingen (optioneel)

Drie plaatsen; `knop-uitlijnen.png` staat er al (jouw schermafbeelding uit les 03). Open
`index.html?leraar` om te zien waar de andere twee komen; de lijst staat in
`assets/screenshots/LEESMIJ.md`. Zolang ze ontbreken, zie je twee 404-meldingen in de console.

---

## 6. Het werkdocument opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Vereist `python-docx`. Het briefje, de wijzigingen en de jobstudenten staan bovenaan het script
(`LEVERINGEN`, `EXTRA_LEVERING`, `JOBSTUDENTEN`). Pas je ze aan, pas dan ook §4 hierboven en rij 1
op de lespagina (stap 2) aan.

---

## 7. Leerplandoelen in je jaaroverzicht

De repository heeft een `pre-push` hook (in `.git/hooks/`, zoals bij les 02–03): bij elke push
roept hij `_tools/update_leerdoelen.py` aan met `lesdoelen.json`. Bij de eerste push op 27-09-2026
zijn de 8 doelen van deze les op het blad *Registratie* van `Leerplandoelen 2026-2027.xlsx`
gezet, en ze staan allemaal op het blad **ORLO**, dus ze worden geteld. Een tweede push voegt niets
dubbel toe.

Met de hand, als dat ooit nodig is:

```bash
python3 "../../_tools/update_leerdoelen.py" lesdoelen.json
```

> [!NOTE]
> De datum in `lesdoelen.json` is een plaatshouder (maandag 28 september 2026). Het script schrijft
> een regel maar één keer: pas je de datum later aan, verbeter hem dan ook op het blad
> *Registratie*. Een hook wordt niet mee gekloond: haal je de repository opnieuw binnen, dan moet
> hij opnieuw geïnstalleerd worden.

## 8. Wat nog moet blijken in de klas

1. **De kolombreedtes na het omzetten.** Het werkdocument zet de vier kolommen van deel A vast op
   even breed. Ik kon niet nagaan of Google Documenten die breedte bij het omzetten overneemt (de
   voorvertoning op de Mac negeert ze). Kijk het na volgens §2. Past alles toch al op één regel,
   dan valt handeling 4 van stap 4 gewoon weg.
2. **Achtergrondkleur in de werkbalk.** Volgens de helppagina staat die knop in de werkbalk als er
   cellen geselecteerd zijn. In een smal venster zit hij waarschijnlijk achter ⋮.
3. **Het raster van *Invoegen › Tabel*.** Kiezen de leerlingen vlot 3 × 6? Zo niet, laat ze een
   kleinere tabel maken en rijen bijvoegen met *Rij onder invoegen*.
4. **Haalbaarheid.** Twee tabellen in 40 minuten keuzewerktijd. Te krap? Gebruik de minimumroute
   uit `lesvoorbereiding.md` §20.
5. **Deadline.** Les 04 en 05 vallen in dezelfde week, met dezelfde deadline (vrijdag 2 oktober,
   20.00 uur). Wie les 05 laat in de week krijgt, heeft na de les weinig tijd om iets af te werken.
6. **Klasnaam.** `lesdoelen.json` gaat uit van `3ORLO`. Pas `klasnaam` aan als dat niet klopt.
