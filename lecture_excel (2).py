"""Lecture des fichiers Excel et normalisation des numéros de téléphone."""
import io
import re
import unicodedata

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
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    chiffres = re.sub(r"\D", "", str(valeur))
    if chiffres.startswith("00"):
        chiffres = chiffres[2:]
    if chiffres.startswith("32") and len(chiffres) >= 10:
        chiffres = chiffres[2:]
    chiffres = chiffres.lstrip("0")
    return chiffres if len(chiffres) >= 7 else ""


def texte_propre(valeur):
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


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


def lire_excel(contenu):
    """Lit un fichier .xlsx (toutes les feuilles).

    Renvoie : (lignes, ignorees, feuilles_ok, feuilles_ko)
    - lignes : liste de (cle, numero_original, nom, prenom, commune)
    - ignorees : nombre de lignes sans numéro valide
    """
    classeur = load_workbook(io.BytesIO(contenu), read_only=True, data_only=True)
    lignes, ignorees, feuilles_ok, feuilles_ko = [], 0, [], []

    for feuille in classeur.worksheets:
        colonnes = None
        titres_vus = []
        try:
            feuille.reset_dimensions()  # certains fichiers annoncent une mauvaise taille
        except Exception:
            pass
        for numero, ligne in enumerate(feuille.iter_rows(values_only=True)):
            if colonnes is None:
                if numero >= 10:
                    break  # pas de ligne de titres trouvée dans les 10 premières
                if numero == 0:
                    titres_vus = [texte_propre(c) for c in ligne if c not in (None, "")]
                types = [type_colonne(c) for c in ligne]
                if "tel" in types and ("nom" in types or "prenom" in types):
                    colonnes = {
                        "tel": [i for i, t in enumerate(types) if t == "tel"],
                        "nom": types.index("nom") if "nom" in types else None,
                        "prenom": types.index("prenom") if "prenom" in types else None,
                        "commune": types.index("commune") if "commune" in types else None,
                    }
                continue

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
                            texte_propre(_case(ligne, colonnes["nom"])),
                            texte_propre(_case(ligne, colonnes["prenom"])),
                            texte_propre(_case(ligne, colonnes["commune"])),
                        )
                    )
                    ajoutees += 1
            if ajoutees == 0:
                ignorees += 1

        if colonnes:
            feuilles_ok.append(feuille.title)
        else:
            feuilles_ko.append(f"{feuille.title} (titres vus : {', '.join(titres_vus[:8]) or 'aucun'})")

    classeur.close()
    return lignes, ignorees, feuilles_ok, feuilles_ko
