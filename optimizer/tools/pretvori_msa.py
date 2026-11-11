"""Pretvara tablicu "MSA zalihe FC" (SAP izvoz s prognozama) u šprancu za StockOptimizer.

Tablica nema mjesečnu prodaju, pa se koristi godišnja potrošnja (aplikacija tada radi bez XYZ analize,
sigurnosne zalihe i narudžbe). Skripta ne sadrži podatke klijenta; ulaz i izlaz zadajete vi.

Upotreba:
    python tools/pretvori_msa.py ulaz.xlsx izlaz.xlsx
    python tools/pretvori_msa.py ulaz.xlsx izlaz.xlsx --potrosnja 2025 --grupa materijal

Mapiranje stupaca (nazivi se traže po sadržaju, pa je svejedno koliko ih ima i kojim redom stoje):
    Šifra            <- Material Number
    Naziv            <- Naziv (ili Material Description)
    Grupa            <- Bakar / Optika  (ili Grupa materijala uz --grupa materijal)
    Nabavna cijena   <- jed.cijena
    Trenutna zaliha  <- Zalihe ukupno (centralno skladište + vanjski partneri + HTS + tranzit)
    Na putu          <- PO neisporučeno kol.
    Lead time (dani) <- Rokovi isporuke / tjedan ("2 do 3" -> 2,5 tjedna), pomnoženo sa 7
    Godišnja potrošnja <- |Razduženo - sve SKLOK| za 1.1.-1.11.2026 sveden na godinu (x 12 / Mjeseci),
                          ili uz --potrosnja 2025 stvarno razduženo u 2025.
"""

import argparse
import re
import sys

import pandas as pd


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s).replace("\n", " ")).strip().lower()


def _nadji(df: pd.DataFrame, *dijelovi: str, tocno: bool = False) -> str | None:
    """Prvi stupac čiji (normalizirani) naziv sadrži sve zadane dijelove."""
    for c in df.columns:
        n = _norm(c)
        if (n == dijelovi[0]) if tocno else all(d in n for d in dijelovi):
            return c
    return None


def tjedni_u_dane(v) -> float | None:
    """'4' -> 28, '2 do 3' -> 17,5, 'do 8' -> 56, prazno ili tekst -> None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    t = str(v).strip().lower().replace(",", ".")
    brojevi = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", t)]
    if not brojevi:
        return None
    return round(sum(brojevi[:2]) / len(brojevi[:2]) * 7, 1)


def pretvori(ulaz: str, potrosnja: str = "2026", grupa: str = "optika") -> pd.DataFrame:
    df = pd.read_excel(ulaz, header=0)
    c_sifra = _nadji(df, "material number")
    if c_sifra is None:
        raise SystemExit("Ne nalazim stupac 'Material Number'. Je li to prava tablica?")
    df = df[df[c_sifra].astype(str).str.strip().str.fullmatch(r"\d+")].copy()   # samo retci s artiklima, bez zbrojeva

    c_naziv = _nadji(df, "naziv", tocno=True) or _nadji(df, "material description")
    c_grupa = (_nadji(df, "bakar / optika") if grupa == "optika" else _nadji(df, "grupa materijala"))
    c_cijena = _nadji(df, "jed.cijena") or _nadji(df, "jed. cijena")
    c_zaliha = _nadji(df, "zalihe ukupno")
    c_po = _nadji(df, "po neisporučeno kol")
    c_lt = _nadji(df, "rokovi isporuke")
    c_mj = _nadji(df, "mjeseci", tocno=True)
    if potrosnja == "2025":
        c_pot = _nadji(df, "razduženo 2025")
    else:
        c_pot = _nadji(df, "razduženo - sve sklok")
    for naziv, c in [("naziv", c_naziv), ("cijena", c_cijena), ("zalihe ukupno", c_zaliha), ("potrošnja", c_pot)]:
        if c is None:
            raise SystemExit(f"Ne nalazim stupac za '{naziv}' u tablici.")

    num = lambda s: pd.to_numeric(s, errors="coerce")
    pot = num(df[c_pot]).abs()
    if potrosnja != "2025":
        mjeseci = num(df[c_mj]).where(lambda m: m > 0).fillna(10) if c_mj else 10
        pot = pot * 12 / mjeseci   # razdoblje 1.1.-1.11. sveden na godinu

    izlaz = pd.DataFrame({
        "Šifra": df[c_sifra].astype(str).str.strip(),
        "Naziv": df[c_naziv].fillna(df[_nadji(df, "material description")] if _nadji(df, "material description") else "").astype(str).str.strip(),
        "Grupa": (df[c_grupa].astype(str).str.strip().str.capitalize() if c_grupa else "(sve)"),
        "Nabavna cijena": num(df[c_cijena]),
        "Trenutna zaliha": num(df[c_zaliha]),
        "Na putu": num(df[c_po]) if c_po else None,
        "Lead time": df[c_lt].map(tjedni_u_dane) if c_lt else None,
        "Godišnja potrošnja": pot.round(1),
    })
    return izlaz.dropna(axis=1, how="all").reset_index(drop=True)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Pretvara MSA tablicu zaliha u šprancu za StockOptimizer.")
    ap.add_argument("ulaz"), ap.add_argument("izlaz")
    ap.add_argument("--potrosnja", choices=["2026", "2025"], default="2026",
                    help="koju potrošnju koristiti: 2026 (1.1.-1.11., svedeno na godinu) ili 2025 (cijela godina)")
    ap.add_argument("--grupa", choices=["optika", "materijal"], default="optika",
                    help="što je 'Grupa': Bakar/Optika ili Grupa materijala")
    a = ap.parse_args(argv)
    t = pretvori(a.ulaz, a.potrosnja, a.grupa)
    t.to_excel(a.izlaz, sheet_name="Podaci", index=False)
    print(f"Spremljeno {len(t)} artikala u {a.izlaz} (stupci: {', '.join(t.columns)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
