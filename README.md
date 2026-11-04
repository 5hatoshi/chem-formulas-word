# Chemické vzorce pro Word

Doplněk do Wordu pro Mac i Windows. Vzorce píšeš obyčejně a on je upraví skutečným dolním a horním indexem.

| Napíšeš | Výsledek |
| --- | --- |
| `H2SO4` | H₂SO₄ |
| `Ca2+` | Ca²⁺ |
| `SO42-` | SO₄²⁻ |
| `CuSO4.5H2O` | CuSO₄·5H₂O |

## Používání

- **Automaticky při psaní:** vzorec se upraví po napsání mezery nebo interpunkce. Vypíná se v panelu Domů → Nastavení vzorců.
- **Ručně:** označ text a stiskni `⌘⇧⌥H` na Macu nebo `Ctrl+Alt+Shift+H` na Windows, případně použij Domů → Upravit vzorce.

Automatika záměrně nechává být běžný text jako A4, MP3, B2B, B12, SN2 nebo O2 Arena. Když označíš jediný výraz a použiješ zkratku, upraví se i B12.

Úprava trvá asi sekundu, protože doplněk si s Wordem vyměňuje data a Word odpovídá s prodlevou. Zkratka nemusí čekat, až si doplněk změny všimne.

## Vývoj

```bash
npm install
```

```bash
npm test
```

```bash
npm run icons
```

```bash
npm run certs
```

```bash
npm run dev
```

`npm run certs` přidá do systému důvěryhodný vývojový certifikát, takže si vyžádá heslo. Stačí to spustit jednou.

## Načtení do Wordu na Macu

Doplněk běží z `https://5hatoshi.github.io/chem-formulas-word/`, takže lokální server není potřeba. Zkopíruj manifest a restartuj Word:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef && cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

## Zveřejnění změn

Po commitu do `main` nahraje obsah složky `addin` na GitHub Pages tenhle příkaz:

```bash
npm run deploy
```

Pro vývoj bez zveřejnění vytvoř `manifest.dev.xml`, načti ho do Wordu místo `manifest.xml` a spusť `npm run dev`:

```bash
npm run manifest:dev
```

## Struktura

- `addin/src/formula.js`: rozpoznání vzorců, bez závislosti na Wordu
- `addin/src/planner.js`: rozhodnutí, co upravit
- `addin/src/word.js`: zápis do dokumentu a automatika
- `addin/taskpane.*`: panel a registrace zkratky
- `manifest.xml`, `addin/shortcuts.json`: popis doplňku pro Word
