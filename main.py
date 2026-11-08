"""Bot Telegram privé « Recherche Wallonie »."""
import asyncio
import hashlib
import logging
import os

import psycopg
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

from lecture_excel import cle_telephone, lire_excel

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("recherche-wallonie")

TOKEN = os.environ["TELEGRAM_TOKEN"]
OWNER_ID = int(os.environ["OWNER_ID"])
DATABASE_URL = os.environ["DATABASE_URL"]


def db():
    return psycopg.connect(DATABASE_URL)


def preparer_base():
    with db() as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS imports (
                id SERIAL PRIMARY KEY,
                nom_fichier TEXT NOT NULL,
                empreinte TEXT NOT NULL UNIQUE,
                nb_lignes INTEGER NOT NULL,
                importe_le TIMESTAMPTZ NOT NULL DEFAULT now()
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS personnes (
                id SERIAL PRIMARY KEY,
                import_id INTEGER NOT NULL REFERENCES imports(id),
                tel_cle TEXT NOT NULL,
                tel_origine TEXT,
                nom TEXT,
                prenom TEXT,
                commune TEXT
            )"""
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_personnes_tel ON personnes (tel_cle)")


def importer(nom_fichier, contenu):
    """Ajoute le fichier aux anciennes données. Ne supprime jamais rien."""
    empreinte = hashlib.sha256(contenu).hexdigest()

    with db() as conn:
        deja = conn.execute(
            "SELECT nom_fichier, importe_le FROM imports WHERE empreinte = %s", (empreinte,)
        ).fetchone()
    if deja:
        return (
            f"⚠️ Ce fichier exact a déjà été importé le {deja[1]:%d/%m/%Y} "
            f"(sous le nom « {deja[0]} »). Je ne l'ajoute pas une 2e fois."
        )

    lignes, ignorees, feuilles_ok, feuilles_ko = lire_excel(contenu)
    if not lignes:
        return (
            "❌ Je n'ai trouvé aucune donnée dans ce fichier.\n"
            "Il faut une ligne de titres avec une colonne pour le téléphone "
            "(Téléphone, Tel, GSM…) et une colonne Nom ou Prénom. "
            "Colonne Commune en plus si possible."
        )

    with db() as conn:
        import_id = conn.execute(
            "INSERT INTO imports (nom_fichier, empreinte, nb_lignes) VALUES (%s, %s, %s) RETURNING id",
            (nom_fichier, empreinte, len(lignes)),
        ).fetchone()[0]
        with conn.cursor() as cur:
            cur.executemany(
                "INSERT INTO personnes (import_id, tel_cle, tel_origine, nom, prenom, commune) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                [(import_id, *ligne) for ligne in lignes],
            )
        total = conn.execute("SELECT count(*) FROM personnes").fetchone()[0]

    texte = f"✅ Fichier importé : {len(lignes)} lignes ajoutées.\n📚 Total dans la base : {total} lignes."
    if ignorees:
        texte += f"\n⚠️ {ignorees} lignes sans numéro valide ont été laissées de côté."
    if feuilles_ko:
        texte += "\n⚠️ Feuilles non lues (titres introuvables) : " + ", ".join(feuilles_ko)
    return texte


def chercher(cle):
    with db() as conn:
        return conn.execute(
            """SELECT p.nom, p.prenom, p.commune, i.nom_fichier, i.importe_le
               FROM personnes p JOIN imports i ON i.id = p.import_id
               WHERE p.tel_cle = %s
               ORDER BY i.importe_le, p.id""",
            (cle,),
        ).fetchall()


def statistiques():
    with db() as conn:
        lignes = conn.execute("SELECT count(*) FROM personnes").fetchone()[0]
        fichiers = conn.execute("SELECT count(*) FROM imports").fetchone()[0]
        dernier = conn.execute("SELECT max(importe_le) FROM imports").fetchone()[0]
    texte = f"📚 {lignes} lignes dans la base\n📁 {fichiers} fichiers importés"
    if dernier:
        texte += f"\n🕒 Dernier import : {dernier:%d/%m/%Y}"
    return texte


async def envoyer_long(update, texte):
    """Telegram limite chaque message : on coupe en morceaux si besoin."""
    morceau = ""
    for ligne in texte.split("\n"):
        if len(morceau) + len(ligne) + 1 > 3800:
            await update.message.reply_text(morceau)
            morceau = ""
        morceau += ligne + "\n"
    if morceau.strip():
        await update.message.reply_text(morceau)


async def commande_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    log.info("Commande /start exécutée : accueil RADAR NOTIFICATIONS")
    await update.message.reply_text("RADAR NOTIFICATIONS")


async def commande_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    numeros = [" ".join(context.args)]
    lignes = [ligne.strip() for ligne in numeros if ligne.strip()]
    if not lignes:
        await update.message.reply_text("Utilisation : /search suivi du numéro de téléphone.")
        return

    reponses = []
    for ligne in lignes[:20]:
        cle = cle_telephone(ligne)
        if not cle:
            reponses.append(f"❓ « {ligne} » n'est pas un numéro valide.")
            continue
        resultats = await asyncio.to_thread(chercher, cle)
        if not resultats:
            reponses.append(f"🔎 {ligne}\nAucun résultat.")
            continue
        morceaux = [f"🔎 {ligne} : {len(resultats)} résultat(s)"]
        for n, (nom, prenom, commune, fichier, date) in enumerate(resultats, 1):
            morceaux.append(
                f"{n}. Nom : {nom or '-'}\n"
                f"    Prénom : {prenom or '-'}\n"
                f"    Commune : {commune or '-'}\n"
                f"    (fichier : {fichier}, {date:%d/%m/%Y})"
            )
        reponses.append("\n".join(morceaux))
    await envoyer_long(update, "\n\n".join(reponses))


async def recevoir_texte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lignes = [l.strip() for l in update.message.text.splitlines() if l.strip()]
    # En privé, seul le propriétaire peut rechercher sans la commande /search.
    reponses = []
    for ligne in lignes[:20]:
        cle = cle_telephone(ligne)
        if not cle:
            reponses.append(f"❓ « {ligne} » n'est pas un numéro valide.")
            continue
        resultats = await asyncio.to_thread(chercher, cle)
        if not resultats:
            reponses.append(f"🔎 {ligne}\nAucun résultat.")
            continue
        morceaux = [f"🔎 {ligne} : {len(resultats)} résultat(s)"]
        for n, (nom, prenom, commune, fichier, date) in enumerate(resultats, 1):
            morceaux.append(
                f"{n}. Nom : {nom or '-'}\n"
                f"    Prénom : {prenom or '-'}\n"
                f"    Commune : {commune or '-'}\n"
                f"    (fichier : {fichier}, {date:%d/%m/%Y})"
            )
        reponses.append("\n".join(morceaux))
    await envoyer_long(update, "\n\n".join(reponses))


async def commande_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(await asyncio.to_thread(statistiques))


async def recevoir_fichier(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    nom = document.file_name or "fichier.xlsx"
    if not nom.lower().endswith(".xlsx"):
        await update.message.reply_text(
            "Je lis seulement les fichiers Excel au format .xlsx. "
            "Dans Excel : Fichier > Enregistrer sous > Classeur Excel (.xlsx)."
        )
        return
    await update.message.reply_text("📥 Fichier reçu, import en cours…")
    try:
        fichier = await document.get_file()
        contenu = bytes(await fichier.download_as_bytearray())
        reponse = await asyncio.to_thread(importer, nom, contenu)
    except Exception:
        log.exception("Erreur pendant l'import")
        reponse = "❌ L'import a échoué. Rien n'a été modifié dans la base. Réessaie, ou envoie-moi le message d'erreur."
    await update.message.reply_text(reponse)


async def au_demarrage(app: Application):
    await asyncio.to_thread(preparer_base)
    log.info("Base prête, bot démarré.")


def main():
    app = Application.builder().token(TOKEN).post_init(au_demarrage).build()

    # /start affiche uniquement le nom du service.
    app.add_handler(CommandHandler("start", commande_start))

    # Recherche par commande accessible à tous, y compris dans les groupes.
    app.add_handler(CommandHandler("search", commande_search))
    app.add_handler(CommandHandler("stats", commande_stats))

    # Seul le propriétaire peut importer des fichiers Excel, en privé ou en groupe.
    proprietaire = filters.User(user_id=OWNER_ID)
    app.add_handler(MessageHandler(proprietaire & filters.Document.ALL, recevoir_fichier))

    # En privé, le propriétaire conserve la recherche par simple envoi du numéro.
    recherche_privee = proprietaire & filters.ChatType.PRIVATE & filters.TEXT & ~filters.COMMAND
    app.add_handler(MessageHandler(recherche_privee, recevoir_texte))

    app.run_polling()


if __name__ == "__main__":
    main()
