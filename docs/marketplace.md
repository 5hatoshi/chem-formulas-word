# Odeslání na Microsoft Marketplace

Podklady pro formulář v Partner Center. Texty pro zveřejnění jsou česky, poznámky pro testery anglicky.

## Název

Chemické vzorce

## Odkazy

- Podpora: https://5hatoshi.github.io/chem-formulas-word/support.html
- Soukromí: https://5hatoshi.github.io/chem-formulas-word/privacy.html
- Manifest k nahrání: `manifest.xml`

## Krátký popis

Píšeš H2SO4 a dostaneš H₂SO₄. Doplněk upravuje vzorce a šipky rovnic při psaní.

## Popis

This application is available in Czech.

Chemické vzorce ušetří ruční nastavování dolních a horních indexů. Vzorec píšeš obyčejně a po mezeře se upraví skutečným formátováním Wordu.

Co umí:

- Anorganické sloučeniny a hydráty: H2SO4 na H₂SO₄, CuSO4.5H2O na CuSO₄·5H₂O
- Ionty: Ca2+ na Ca²⁺, SO42- na SO₄²⁻
- Organické sloučeniny a obecné vzorce: CH3COOH na CH₃COOH, CnH2n+2 na CₙH₂ₙ₊₂
- Komplexy včetně ligandů: [Cu(NH3)4]2+ na [Cu(NH₃)₄]²⁺, [Ni(en)3]2+ na [Ni(en)₃]²⁺
- Šipky rovnic: -> na →, <=> na ⇌, CO2^ na CO₂↑, BaSO4v na BaSO₄↓

Běžný text jako A4, MP3 nebo B2B nechává být. Automatiku lze vypnout v panelu a každou úpravu vrátit pomocí Zpět. V novém dokumentu se automatika spustí po prvním otevření doplňku.

Doplněk je zdarma, nevyžaduje přihlášení a obsah dokumentů nikam neposílá.

## Klíčová slova

chemie, vzorce, dolní index, horní index, chemistry

## Poznámky pro testery (Notes for certification)

No account, sign-in or sample data is needed. The UI is in Czech.

1. Open the add-in with Home > "Nastavení vzorců". This opens the task pane. Automatic conversion ("Upravovat automaticky při psaní") is on by default and can be switched off there.
2. Type `H2SO4 Ca2+ SO42- CuSO4.5H2O` followed by a space. Expected: H₂SO₄ Ca²⁺ SO₄²⁻ CuSO₄·5H₂O. Conversion happens about a second after the space.
3. Type `CaCO3 -> CaO + CO2^` followed by a space. Expected: CaCO₃ → CaO + CO₂↑.
4. Press Undo after a conversion. The formula stays as typed.
5. Switch automatic conversion off, type `H2O`, select it and click "Upravit výběr" in the task pane. This covers touch-only devices. The ribbon button "Upravit vzorce" does the same. The keyboard shortcut is Ctrl+Alt+Shift+H, or Cmd+Shift+Option+H on Mac.

The add-in changes the document only to format formulas. Automatic conversion is described in the task pane, can be switched off and can be undone. It sends no data anywhere.

## Co musíš udělat sám

1. Vytvořit účet v [Partner Center](https://partner.microsoft.com/dashboard) a projít ověřením. Ověř, jaké údaje a případné poplatky to vyžaduje, Microsoft mluví o firemním účtu.
2. Pořídit snímky obrazovky se zobrazeným doplňkem ve Wordu. Požadované rozměry jsou uvedené přímo ve formuláři.
3. Založit nabídku typu Office Add-in, nahrát `manifest.xml` a vyplnit texty výše.
4. Odeslat ke kontrole. Podle Microsoftu trvá 3 až 5 pracovních dní.

Po schválení se každá změna manifestu musí odeslat znovu a s vyšším číslem verze.
