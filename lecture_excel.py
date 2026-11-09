"""Import Excel sans en-têtes : A=genre, B=nom, C=prénom, reste détecté automatiquement."""
import difflib
import io
import json
import re
import unicodedata
import urllib.request

from openpyxl import load_workbook

_API_COMMUNES = "https://www.odwb.be/api/explore/v2.1/catalog/datasets/communes_s3/records?limit=300"
_COMMUNES = None


def simplifier(texte):
    t = unicodedata.normalize("NFD", str(texte or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t.lower())


def texte_propre(valeur):
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return re.sub(r"\s+", " ", str(valeur)).strip()


def cle_telephone(valeur):
    if valeur is None or isinstance(valeur, bool):
        return ""
    brut = str(int(valeur)) if isinstance(valeur, float) and valeur.is_integer() else str(valeur)
    chiffres = re.sub(r"\D", "", brut)
    if chiffres.startswith("00"):
        chiffres = chiffres[2:]
    if chiffres.startswith("32") and len(chiffres) >= 10:
        chiffres = chiffres[2:]
    chiffres = chiffres.lstrip("0")
    return chiffres if 7 <= len(chiffres) <= 10 else ""


def _telephone_valide(valeur):
    return bool(cle_telephone(valeur))


def _charger_communes():
    """Charge le référentiel public UVCW/ODWB; garde le bot utilisable si l'API est indisponible."""
    global _COMMUNES
    if _COMMUNES is not None:
        return _COMMUNES
    _COMMUNES = {}
    try:
        req = urllib.request.Request(_API_COMMUNES, headers={"User-Agent": "recherche-wallonie/1.0"})
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode("utf-8"))
        for record in data.get("results", []):
            nom = texte_propre(record.get("nom"))
            court = texte_propre(record.get("nom_court"))
            if nom:
                _COMMUNES[simplifier(nom)] = nom
            if court:
                _COMMUNES[simplifier(court)] = nom or court
    except Exception:
        # Pas de correction hasardeuse si la source officielle est indisponible.
        pass
    return _COMMUNES


def corriger_commune(valeur):
    brut = texte_propre(valeur)
    if not brut:
        return ""
    communes = _charger_communes()
    cle = simplifier(brut)
    if not communes:
        return brut
    if cle in communes:
        return communes[cle]
    if len(cle) < 4:
        return brut
    candidat = difflib.get_close_matches(cle, communes.keys(), n=1, cutoff=0.84)
    return communes[candidat[0]] if candidat else brut


def _case(ligne, index):
    return ligne[index] if index is not None and index < len(ligne) else None


def _genre(valeur):
    return texte_propre(valeur)


def _inferer_colonnes_supplementaires(lignes):
    """Repère le téléphone et la commune dans D et les colonnes suivantes."""
    largeur = max((len(r) for r in lignes), default=0)
    telephone, commune = None, None
    meilleur_tel, meilleur_commune = 0.0, 0.0
    communes = _charger_communes()

    for index in range(3, largeur):
        valeurs = [r[index] for r in lignes if index < len(r) and texte_propre(r[index])]
        if not valeurs:
            continue
        score_tel = sum(_telephone_valide(v) for v in valeurs) / len(valeurs)
        if score_tel >= 0.55 and score_tel > meilleur_tel:
            telephone, meilleur_tel = index, score_tel

        if communes:
            score_commune = sum(
                1 for v in valeurs
                if simplifier(v) in communes or (
                    len(simplifier(v)) >= 4 and
                    bool(difflib.get_close_matches(simplifier(v), communes.keys(), n=1, cutoff=0.84))
                )
            ) / len(valeurs)
            if score_commune >= 0.50 and score_commune > meilleur_commune:
                commune, meilleur_commune = index, score_commune

    return telephone, commune


def lire_excel(contenu):
    """Renvoie (lignes, ignorees, feuilles_ok, feuilles_ko).
    Chaque ligne importée vaut (clé_tel, tel_origine, genre, nom, prénom, commune).
    Aucune ligne d'en-tête n'est attendue : A/B/C sont toujours genre/nom/prénom.
    """
    classeur = load_workbook(io.BytesIO(contenu), read_only=True, data_only=True)
    lignes, ignorees, feuilles_ok, feuilles_ko = [], 0, [], []

    try:
        for feuille in classeur.worksheets:
            donnees = [
                tuple(texte_propre(v) if v is not None else "" for v in row)
                for row in feuille.iter_rows(values_only=True)
            ]
            donnees = [r for r in donnees if any(v != "" for v in r)]
            if not donnees:
                feuilles_ko.append(feuille.title)
                continue

            tel_col, commune_col = _inferer_colonnes_supplementaires(donnees[:min(200, len(donnees))])
            if tel_col is None:
                feuilles_ko.append(feuille.title)
                continue

            feuilles_ok.append(feuille.title)
            for row in donnees:
                brut_tel = _case(row, tel_col)
                cle = cle_telephone(brut_tel)
                if not cle:
                    ignorees += 1
                    continue
                genre = _genre(_case(row, 0))
                nom = texte_propre(_case(row, 1))
                prenom = texte_propre(_case(row, 2))
                commune = corriger_commune(_case(row, commune_col)) if commune_col is not None else ""
                lignes.append((cle, texte_propre(brut_tel), genre, nom, prenom, commune))
    finally:
        classeur.close()

    return lignes, ignorees, feuilles_ok, feuilles_ko
