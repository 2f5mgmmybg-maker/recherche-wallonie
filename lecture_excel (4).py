"""Lecture des fichiers Excel et normalisation des numéros de téléphone."""
import io
import re
import unicodedata
from datetime import date, datetime

from openpyxl import load_workbook


def simplifier(texte):
    """Minuscules, sans accents, sans espaces ni signes."""
    t = unicodedata.normalize("NFD", str(texte))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t.lower())


def cle_telephone(valeur):
    """Transforme un numéro en clé de recherche unique.

    0475 12 34 56, +32 475 12 34 56, 0032475123456 et 475123456
    donnent tous la même clé : 475123456.
    Renvoie "" si ce n'est pas un numéro valide.
    """
    if valeur is None or isinstance(valeur, (date, datetime)):
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    if re.match(r"^\s*\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\s*$", str(valeur)):
        return ""  # c'est une date, pas un numéro
    chiffres = re.sub(r"\D", "", str(valeur))
    if chiffres.startswith("00"):
        chiffres = chiffres[2:]
    if chiffres.startswith("32") and len(chiffres) >= 10:
        chiffres = chiffres[2:]
    chiffres = chiffres.lstrip("0")
    return chiffres if 7 <= len(chiffres) <= 14 else ""


def texte_propre(valeur):
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


def reparer_texte(valeur):
    """Répare un texte abîmé du genre  b'Li\\xc3\\xa8ge, Belgium'  (devient : Liège, Belgium)."""
    t = texte_propre(valeur)
    m = re.match(r"^b(['\"])(.*)\1$", t, re.S)
    if m:
        t = m.group(2)
    t = re.sub(
        r"(?:\\+x[0-9a-fA-F]{2})+",
        lambda x: bytes.fromhex(re.sub(r"\\+x", "", x.group(0))).decode("utf-8", "replace"),
        t,
    )
    return t.strip()


def nettoyer_commune(valeur):
    """Garde seulement la commune : « Liège, Belgium » devient « Liège »."""
    t = reparer_texte(valeur)
    return t.split(",")[0].strip() if t else ""


def type_colonne(entete):
    """Devine ce que contient une colonne d'après son titre."""
    h = simplifier(entete) if entete is not None else ""
    if not h:
        return None
    if "prenom" in h or "firstname" in h:
        return "prenom"
    if any(m in h for m in ("commune", "ville", "localite", "city", "municipalite")):
        return "commune"
    if any(m in h for m in ("tel", "gsm", "mobile", "portable", "phone")) or h in (
        "numero",
        "num",
        "number",
    ):
        return "tel"
    if h.startswith("nom") or h in ("lastname", "surname", "name"):
        return "nom"
    return None


def _case(ligne, i):
    if i is None or i >= len(ligne):
        return None
    return ligne[i]


def _colonnes_avec_titres(lignes):
    """Cherche une ligne de titres dans les 10 premières lignes."""
    for numero, ligne in enumerate(lignes[:10]):
        types = [type_colonne(c) for c in ligne]
        if "tel" in types and ("nom" in types or "prenom" in types):
            return numero, {
                "tel": [i for i, t in enumerate(types) if t == "tel"],
                "nom": types.index("nom") if "nom" in types else None,
                "prenom": types.index("prenom") if "prenom" in types else None,
                "commune": types.index("commune") if "commune" in types else None,
            }
    return None, None


def _colonnes_sans_titres(lignes):
    """Fichier sans titres. Ordre prévu : Femme / Prénom / Nom / Numéro / Ville / Date de naissance.

    On repère d'abord la colonne qui contient des numéros (la 4e en priorité),
    puis Prénom = 2 colonnes avant, Nom = 1 avant, Commune = 1 après.
    """
    echantillon = [l for l in lignes if any(c not in (None, "") for c in l)][:200]
    if not echantillon:
        return None
    largeur = max(len(l) for l in echantillon)
    meilleur, score = None, 0.0
    for i in range(largeur):
        bons = sum(1 for l in echantillon if cle_telephone(_case(l, i)))
        ratio = bons / len(echantillon)
        # la 4e colonne (index 3) gagne en cas d'égalité
        if ratio > score + 1e-9 or (abs(ratio - score) < 1e-9 and i == 3 and ratio > 0):
            meilleur, score = i, ratio
    if meilleur is None or score < 0.5:
        return None
    return {
        "tel": [meilleur],
        "prenom": meilleur - 2 if meilleur >= 2 else None,
        "nom": meilleur - 1 if meilleur >= 1 else None,
        "commune": meilleur + 1,
    }


def lire_excel(contenu):
    """Lit un fichier .xlsx (toutes les feuilles), avec ou sans ligne de titres.

    Renvoie : (lignes, ignorees, feuilles_ok, feuilles_ko)
    - lignes : liste de (cle, numero_original, nom, prenom, commune)
    - ignorees : nombre de lignes sans numéro valide
    """
    classeur = load_workbook(io.BytesIO(contenu), read_only=True, data_only=True)
    lignes, ignorees, feuilles_ok, feuilles_ko = [], 0, [], []

    for feuille in classeur.worksheets:
        try:
            feuille.reset_dimensions()  # certains fichiers annoncent une mauvaise taille
        except Exception:
            pass
        toutes = list(feuille.iter_rows(values_only=True))

        debut, colonnes = _colonnes_avec_titres(toutes)
        if colonnes is not None:
            donnees = toutes[debut + 1 :]
        else:
            colonnes = _colonnes_sans_titres(toutes)
            donnees = toutes

        if colonnes is None:
            vus = [texte_propre(c) for c in (toutes[0] if toutes else []) if c not in (None, "")]
            feuilles_ko.append(f"{feuille.title} (1re ligne : {', '.join(vus[:8]) or 'vide'})")
            continue

        for ligne in donnees:
            if not any(c not in (None, "") for c in ligne):
                continue  # ligne vide
            ajoutees = 0
            for i in colonnes["tel"]:
                brut = _case(ligne, i)
                cle = cle_telephone(brut)
                if cle:
                    lignes.append(
                        (
                            cle,
                            texte_propre(brut),
                            reparer_texte(_case(ligne, colonnes["nom"])),
                            reparer_texte(_case(ligne, colonnes["prenom"])),
                            nettoyer_commune(_case(ligne, colonnes["commune"])),
                        )
                    )
                    ajoutees += 1
            if ajoutees == 0:
                ignorees += 1
        feuilles_ok.append(feuille.title)

    classeur.close()
    return lignes, ignorees, feuilles_ok, feuilles_ko
