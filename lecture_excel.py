"""Lecture Excel avec détection des colonnes par leur contenu, même sans en-têtes."""
import io
import re
import unicodedata

from openpyxl import load_workbook


def simplifier(texte):
    t = unicodedata.normalize("NFD", str(texte))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", t.lower())


def cle_telephone(valeur):
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
    return chiffres if 7 <= len(chiffres) <= 10 else ""


def texte_propre(valeur):
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


def type_colonne(entete):
    h = simplifier(entete) if entete is not None else ""
    if not h:
        return None
    if "prenom" in h or "firstname" in h or h == "first":
        return "prenom"
    if any(m in h for m in ("commune", "ville", "localite", "city", "municipalite", "postal", "codepostal")):
        return "commune"
    if any(m in h for m in ("tel", "gsm", "mobile", "portable", "phone")) or h in ("numero", "num", "number"):
        return "tel"
    if h.startswith("nom") or h in ("lastname", "surname", "name", "nomdefamille"):
        return "nom"
    return None


def _case(ligne, i):
    return ligne[i] if i is not None and i < len(ligne) else None


def _telephone_valide(valeur):
    if valeur is None or isinstance(valeur, bool):
        return False
    brut = str(int(valeur)) if isinstance(valeur, float) and valeur.is_integer() else str(valeur)
    chiffres = re.sub(r"\D", "", brut)
    if chiffres.startswith("00"):
        chiffres = chiffres[2:]
    if chiffres.startswith("32"):
        chiffres = chiffres[2:]
    chiffres = chiffres.lstrip("0")
    return 7 <= len(chiffres) <= 10


def _texte_humain(valeur):
    if not isinstance(valeur, str):
        return False
    v = valeur.strip()
    if len(v) < 2 or len(v) > 60 or re.search(r"\d", v):
        return False
    mots = re.findall(r"[A-Za-zÀ-ÿ'’-]+", v)
    return 1 <= len(mots) <= 4 and len("".join(mots)) >= 2


def _inferer_colonnes(lignes_brutes):
    """Retourne les colonnes probables par examen des valeurs, sans supposer leur ordre."""
    if not lignes_brutes:
        return None

    largeur = max((len(r) for r in lignes_brutes), default=0)
    telephone_scores = []
    textes = []
    for i in range(largeur):
        valeurs = [r[i] for r in lignes_brutes if i < len(r) and r[i] not in (None, "")]
        if not valeurs:
            continue
        tel = sum(_telephone_valide(v) for v in valeurs)
        if tel >= 1 and tel / len(valeurs) >= 0.55:
            telephone_scores.append((tel / len(valeurs), tel, i))
        humains = [v for v in valeurs if _texte_humain(v)]
        if len(humains) >= 2 and len(humains) / len(valeurs) >= 0.65:
            textes.append((len(humains) / len(valeurs), len(humains), i, humains))

    if not telephone_scores:
        return None
    telephone_scores.sort(reverse=True)
    tel_col = telephone_scores[0][2]

    # Les champs texte sont facultatifs. On ne les attribue que si leur contenu
    # est clairement textuel ; en cas de doute, les numéros restent importables.
    textes = [t for t in textes if t[2] != tel_col]
    colonnes = {"tel": [tel_col], "nom": None, "prenom": None, "commune": None}

    # Détecter les communes grâce à des codes postaux belges lorsqu'ils existent
    # dans une colonne voisine : la valeur texte associée est alors la commune.
    codes_postaux = set()
    for r in lignes_brutes:
        for v in r:
            s = str(v).strip() if v is not None else ""
            if re.fullmatch(r"\d{4}", s) and 1000 <= int(s) <= 9999:
                codes_postaux.add(s)
    # Sans dictionnaire géographique ni en-tête, une commune et un nom peuvent
    # être indiscernables : conserver les colonnes textuelles restantes sans
    # inventer de correspondance sémantique.
    if len(textes) == 1:
        colonnes["nom"] = textes[0][2]
    elif len(textes) >= 2:
        # Le prénom est généralement une seule courte expression ; le nom peut
        # comporter plusieurs mots. Si le profil ne les distingue pas, garder
        # le premier champ comme nom et le second comme prénom.
        profils = []
        for score, count, idx, vals in textes:
            moyenne_mots = sum(len(re.findall(r"[A-Za-zÀ-ÿ'’-]+", v)) for v in vals) / len(vals)
            moyenne_longueur = sum(len(v.strip()) for v in vals) / len(vals)
            profils.append((moyenne_mots, moyenne_longueur, idx))
        profils.sort()
        colonnes["prenom"] = profils[0][2]
        colonnes["nom"] = profils[1][2]
    return colonnes


def lire_excel(contenu):
    """Lit toutes les feuilles .xlsx, avec ou sans ligne de titres.

    Renvoie (lignes, ignorees, feuilles_ok, feuilles_ko).
    Une feuille sans titres est analysée par le contenu des cellules.
    """
    classeur = load_workbook(io.BytesIO(contenu), read_only=True, data_only=True)
    lignes, ignorees, feuilles_ok, feuilles_ko = [], 0, [], []

    for feuille in classeur.worksheets:
        donnees = [tuple(texte_propre(v) if v is not None else None for v in row)
                   for row in feuille.iter_rows(values_only=True)]
        donnees = [r for r in donnees if any(v not in (None, "") for v in r)]
        if not donnees:
            feuilles_ko.append(feuille.title)
            continue

        colonnes = None
        debut = 0
        # Chercher des en-têtes dans les dix premières lignes, sans exiger
        # qu'ils soient sur la première ligne.
        for pos, row in enumerate(donnees[:10]):
            types = [type_colonne(v) for v in row]
            if "tel" in types and ("nom" in types or "prenom" in types):
                colonnes = {
                    "tel": [i for i, t in enumerate(types) if t == "tel"],
                    "nom": types.index("nom") if "nom" in types else None,
                    "prenom": types.index("prenom") if "prenom" in types else None,
                    "commune": types.index("commune") if "commune" in types else None,
                }
                debut = pos + 1
                break

        if colonnes is None:
            # Examiner un échantillon de lignes, puis traiter toutes les lignes
            # comme des données : aucune ligne n'est supprimée comme titre.
            colonnes = _inferer_colonnes(donnees[:min(len(donnees), 100)])
            debut = 0

        if colonnes is None:
            feuilles_ko.append(feuille.title)
            continue

        feuilles_ok.append(feuille.title)
        for row in donnees[debut:]:
            ajoutees = 0
            for i in colonnes["tel"]:
                brut = _case(row, i)
                cle = cle_telephone(brut)
                if cle:
                    lignes.append((
                        cle,
                        texte_propre(brut),
                        texte_propre(_case(row, colonnes["nom"])),
                        texte_propre(_case(row, colonnes["prenom"])),
                        texte_propre(_case(row, colonnes["commune"])),
                    ))
                    ajoutees += 1
            if not ajoutees:
                ignorees += 1

    classeur.close()
    return lignes, ignorees, feuilles_ok, feuilles_ko
