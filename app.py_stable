import io
import os
import time
import json
import random
import requests
import pandas as pd
import streamlit as st
from PIL import Image
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment


# ==========================
# CONFIG GLOBALE
# ==========================

IMAGE_COLS = ["image_url_1", "image_url_2", "image_url_3"]
SERPAPI_ENDPOINT = "https://serpapi.com/search.json"
CACHE_PATH = "image_cache.json"

# Login / mot de passe via variables d'environnement
VALID_USERNAME = os.getenv("APP_USERNAME", "admin")
VALID_PASSWORD = os.getenv("APP_PASSWORD", "change_me")

# Clés SerpAPI pré-configurées (liste séparée par des virgules)
# Exemple: SERPAPI_KEYS="cle1,cle2,cle3"
SERPAPI_KEYS_RAW = os.getenv("SERPAPI_KEYS", "").strip()


# ==========================
# AUTH + CAPTCHA
# ==========================

def init_session_state():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "captcha_question" not in st.session_state:
        regenerate_captcha()


def regenerate_captcha():
    a = random.randint(1, 9)
    b = random.randint(1, 9)
    st.session_state.captcha_question = f"{a} + {b}"
    st.session_state.captcha_expected = a + b


def login_form():
    st.title("🔐 Accès sécurisé au générateur de catalogue d'images")
    st.write("Merci de vous authentifier pour utiliser l’outil.")

    username = st.text_input("Identifiant")
    password = st.text_input("Mot de passe", type="password")
    st.write(f"Captcha : **{st.session_state.captcha_question} = ?**")
    captcha_input = st.text_input("Réponse au captcha")

    if st.button("Se connecter"):
        ok_creds = (username == VALID_USERNAME and password == VALID_PASSWORD)
        ok_captcha = (
            captcha_input.strip().isdigit()
            and int(captcha_input.strip()) == st.session_state.captcha_expected
        )

        if ok_creds and ok_captcha:
            st.session_state.authenticated = True
            st.success("Connexion réussie ✅")
        else:
            st.error("Identifiants ou captcha incorrects.")
            regenerate_captcha()


# ==========================
# CACHE GLOBAL (fichier JSON)
# ==========================

def load_cache(cache_path: str = CACHE_PATH):
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(cache: dict, cache_path: str = CACHE_PATH):
    try:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        # on ignore les erreurs de sauvegarde pour ne pas bloquer l'app
        pass


# ==========================
# SERPAPI KEY SELECTION
# ==========================

def get_configured_serpapi_keys():
    """
    Retourne une liste de clés issues de SERPAPI_KEYS (env),
    ex: SERPAPI_KEYS="key1,key2,key3"
    """
    if not SERPAPI_KEYS_RAW:
        return []
    keys = [k.strip() for k in SERPAPI_KEYS_RAW.split(",") if k.strip()]
    return keys


def serpapi_key_selector():
    """
    UI pour sélectionner une clé SerpAPI :
    - soit parmi les clés configurées dans SERPAPI_KEYS
    - soit en saisissant une nouvelle clé manuellement
    """
    configured_keys = get_configured_serpapi_keys()

    serpapi_key = None

    if configured_keys:
        options = []
        for i, key in enumerate(configured_keys):
            last4 = key[-4:] if len(key) >= 4 else key
            options.append(f"Clé {i+1} (****{last4})")

        options.append("Saisir une autre clé...")

        choice = st.selectbox(
            "Choisir une clé SerpAPI parmi la configuration ou saisir une nouvelle clé :",
            options=options,
        )

        if choice == "Saisir une autre clé...":
            serpapi_key = st.text_input(
                "Nouvelle clé SerpAPI",
                type="password",
                help="Cette clé n'est pas stockée, elle est utilisée uniquement pour cette session.",
            )
        else:
            idx = options.index(choice)
            serpapi_key = configured_keys[idx]
            st.info(f"Clé SerpAPI sélectionnée : {choice}")
    else:
        serpapi_key = st.text_input(
            "Clé API SerpAPI",
            type="password",
            help="Aucune clé préconfigurée. Crée une clé sur https://serpapi.com/ puis colle-la ici.",
        )

    return serpapi_key


# ==========================
# RECHERCHE D’IMAGES
# ==========================

def build_query(row: pd.Series, search_columns):
    parts = []
    for col in search_columns:
        value = row.get(col)
        if pd.notna(value):
            text = str(value).strip()
            if text:
                parts.append(text)
    return " ".join(parts)


def search_images_serpapi(query: str, serpapi_key: str, max_results: int = 3):
    params = {
        "engine": "google_images",
        "q": query,
        "num": max_results,
        "api_key": serpapi_key,
    }

    response = requests.get(SERPAPI_ENDPOINT, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()

    results = data.get("images_results", [])
    urls = []
    for r in results:
        url = r.get("original") or r.get("thumbnail")
        if url:
            urls.append(url)
        if len(urls) >= max_results:
            break

    return urls


def enrich_df_with_images(
    df: pd.DataFrame,
    serpapi_key: str,
    search_columns,
    max_images: int = 3,
    sleep_seconds: float = 0.2,
    test_mode: bool = False,
    max_rows_test: int = 50,
    use_global_cache: bool = True,
    use_unique_key_cache: bool = True,
    unique_key_columns=None,
):
    """
    Enrichit un DataFrame avec des URLs d'images.

    - use_global_cache : utilise un cache texte->images (image_cache.json).
    - use_unique_key_cache : réutilise les images pour les mêmes clés uniques (REF, REF+EAN, etc.).
    - unique_key_columns : liste de colonnes composant la clé unique (facultatif).
    """
    if not search_columns:
        st.error("Aucune colonne sélectionnée pour la recherche d'images.")
        return df, []

    if unique_key_columns is None:
        unique_key_columns = []

    # Colonnes de sortie
    for i, col in enumerate(IMAGE_COLS[:max_images]):
        if col not in df.columns:
            df[col] = None

    # Chargement du cache global
    global_cache = load_cache() if use_global_cache else {}
    # Cache clé unique : clé composite -> liste d'URLs
    unique_cache = {} if use_unique_key_cache and unique_key_columns else {}

    log_entries = []

    total_rows = len(df)
    rows_to_process = min(max_rows_test, total_rows) if test_mode else total_rows

    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, (row_index, row) in enumerate(df.iterrows(), start=1):
        if test_mode and idx > rows_to_process:
            break

        query = build_query(row, search_columns).strip()

        # Construire la clé unique (composite) si configurée
        unique_key = None
        if use_unique_key_cache and unique_key_columns:
            unique_key = tuple(str(row.get(col, "")).strip() for col in unique_key_columns)

        urls = []
        source = None  # "unique_key_cache" / "global_cache" / "serpapi" / None

        if not query:
            status_text.text(f"Ligne {row_index}: requête vide, ignorée.")
        else:
            # 1) Cache par clé unique (super optimisé)
            if unique_key is not None and unique_key in unique_cache:
                urls = unique_cache[unique_key]
                source = "unique_key_cache"

            else:
                # 2) Cache global texte -> URLs
                if use_global_cache and query in global_cache:
                    urls = global_cache[query]
                    source = "global_cache"
                else:
                    # 3) Appel SerpAPI
                    try:
                        urls = search_images_serpapi(
                            query=query,
                            serpapi_key=serpapi_key,
                            max_results=max_images,
                        )
                        source = "serpapi"
                        if use_global_cache:
                            global_cache[query] = urls
                        if not test_mode:
                            time.sleep(sleep_seconds)
                    except Exception as e:
                        status_text.text(f"Ligne {row_index}: ERREUR ({e})")
                        urls = []

                # On alimente le cache clé unique, même si la liste est vide
                if unique_key is not None and use_unique_key_cache:
                    unique_cache[unique_key] = urls

            # Remplissage des colonnes image_url_X
            for i, url in enumerate(urls):
                df.at[row_index, IMAGE_COLS[i]] = url

            status_text.text(
                f"Ligne {row_index}: {len(urls)} image(s) "
                f"{'via ' + source if source else ''} pour '{query[:60]}...'"
            )

            log_entries.append(
                {
                    "row_index": row_index,
                    "query": query,
                    "source": source,
                    "unique_key": unique_key,
                    "num_images": len(urls),
                    "urls": urls,
                }
            )

        progress_bar.progress(idx / rows_to_process)

    status_text.text("Recherche d'images terminée ✅")

    # Sauvegarde du cache global
    if use_global_cache:
        save_cache(global_cache)

    return df, log_entries




# ==========================
# EXCEL + IMAGES
# ==========================
def create_excel_with_embedded_images(
    df: pd.DataFrame,
    image_cols,
    jpg_quality: int = 80,
    max_width: int = 80,
    max_height: int = 80,
    unique_key_columns=None,
):
    """
    Crée un fichier Excel en mémoire avec :
    - les données du DataFrame (toutes les lignes),
    - les images téléchargées, converties en JPG,
      insérées dans les cellules correspondant aux colonnes image_cols,
      en conservant les proportions d'origine,
    - si unique_key_columns est défini :
        * les images sont insérées une seule fois par combinaison de clé
          (sur la première ligne du groupe),
        * les cellules des colonnes de la clé + colonnes photos sont mergées
          verticalement sur les lignes du groupe.
    - formattage : contenu centré (H/V) + largeur colonnes ajustée + hauteur
      augmentée pour les lignes avec image.
    """
    if unique_key_columns is None:
        unique_key_columns = []

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Écrire les données brutes
        df.to_excel(writer, index=False, sheet_name="Produits")
        wb = writer.book
        ws = wb["Produits"]

        rows_with_image = set()

        # 2) Gestion des groupes par clé unique (si clé définie)
        groups = []
        if unique_key_columns:
            current_key = None
            start_excel_row = None
            prev_excel_row = None

            for row_idx, row in df.iterrows():
                excel_row = row_idx + 2
                key = tuple(str(row.get(col, "")).strip() for col in unique_key_columns)

                if key != current_key:
                    if current_key is not None and start_excel_row is not None:
                        groups.append((current_key, start_excel_row, prev_excel_row))
                    current_key = key
                    start_excel_row = excel_row

                prev_excel_row = excel_row

            if current_key is not None and start_excel_row is not None:
                groups.append((current_key, start_excel_row, prev_excel_row))

        # 3) Insertion des images
        if unique_key_columns and groups:
            # Une fois par groupe, sur la première ligne du groupe
            for key, start_row, end_row in groups:
                # Index pandas correspondant à la première ligne du groupe
                df_idx = start_row - 2
                if df_idx < 0 or df_idx >= len(df):
                    continue
                row = df.iloc[df_idx]

                for img_col in image_cols:
                    if img_col not in df.columns:
                        continue

                    url = row.get(img_col)
                    if not isinstance(url, str) or not url.strip():
                        continue

                    try:
                        resp = requests.get(url, timeout=30)
                        resp.raise_for_status()

                        pil_img = Image.open(io.BytesIO(resp.content))
                        if pil_img.mode != "RGB":
                            pil_img = pil_img.convert("RGB")

                        orig_w, orig_h = pil_img.size

                        # Respecter max_width / max_height sans agrandir
                        scale_w = max_width / orig_w
                        scale_h = max_height / orig_h
                        scale = min(scale_w, scale_h, 1.0)

                        new_w = int(orig_w * scale)
                        new_h = int(orig_h * scale)

                        img_bytes = io.BytesIO()
                        pil_img.save(
                            img_bytes,
                            format="JPEG",
                            quality=jpg_quality,
                            optimize=True,
                        )
                        img_bytes.seek(0)

                        xl_img = XLImage(img_bytes)
                        xl_img.width = new_w
                        xl_img.height = new_h

                        col_idx = df.columns.get_loc(img_col) + 1
                        col_letter = get_column_letter(col_idx)
                        cell_ref = f"{col_letter}{start_row}"  # une seule fois, sur la 1re ligne

                        ws.add_image(xl_img, cell_ref)
                        rows_with_image.add(start_row)

                    except Exception:
                        continue
        else:
            # Pas de clé unique : on met les images ligne par ligne
            for row_idx, row in df.iterrows():
                excel_row = row_idx + 2

                for img_col in image_cols:
                    if img_col not in df.columns:
                        continue

                    url = row.get(img_col)
                    if not isinstance(url, str) or not url.strip():
                        continue

                    try:
                        resp = requests.get(url, timeout=30)
                        resp.raise_for_status()

                        pil_img = Image.open(io.BytesIO(resp.content))
                        if pil_img.mode != "RGB":
                            pil_img = pil_img.convert("RGB")

                        orig_w, orig_h = pil_img.size

                        scale_w = max_width / orig_w
                        scale_h = max_height / orig_h
                        scale = min(scale_w, scale_h, 1.0)

                        new_w = int(orig_w * scale)
                        new_h = int(orig_h * scale)

                        img_bytes = io.BytesIO()
                        pil_img.save(
                            img_bytes,
                            format="JPEG",
                            quality=jpg_quality,
                            optimize=True,
                        )
                        img_bytes.seek(0)

                        xl_img = XLImage(img_bytes)
                        xl_img.width = new_w
                        xl_img.height = new_h

                        col_idx = df.columns.get_loc(img_col) + 1
                        col_letter = get_column_letter(col_idx)
                        cell_ref = f"{col_letter}{excel_row}"

                        ws.add_image(xl_img, cell_ref)
                        rows_with_image.add(excel_row)

                    except Exception:
                        continue

        # 4) Merge des cellules pour les groupes de même clé (sur les colonnes clé + colonnes photos)
        if unique_key_columns and groups:
            cols_to_merge = list(unique_key_columns)
            cols_to_merge.extend(image_cols)  # on merge toutes les colonnes photo

            for key, start_row, end_row in groups:
                if end_row <= start_row:
                    continue  # une seule ligne, rien à merger

                for col_name in cols_to_merge:
                    if col_name not in df.columns:
                        continue
                    col_idx = df.columns.get_loc(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    cell_range = f"{col_letter}{start_row}:{col_letter}{end_row}"
                    ws.merge_cells(cell_range)

        # 5) Mise en forme : centrage H/V + wrap + largeur colonnes + hauteur lignes
        # Alignement pour toutes les cellules du tableau
        for row in ws.iter_rows(
            min_row=1, max_row=ws.max_row,
            min_col=1, max_col=ws.max_column
        ):
            for cell in row:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

        # Largeur des colonnes en fonction du contenu
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_length = 0
            for cell in ws[col_letter]:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            # limite pour éviter des colonnes gigantesques
            ws.column_dimensions[col_letter].width = min(max_length * 1.2 + 2, 50)

        # Hauteur des lignes : plus grande pour celles qui ont une image
        base_height = 18
        img_height = max(max_height, 60)  # au moins 60 pour bien voir la chaussure etc.
        for row_idx in range(1, ws.max_row + 1):
            if row_idx in rows_with_image:
                ws.row_dimensions[row_idx].height = img_height
            else:
                ws.row_dimensions[row_idx].height = base_height

    output.seek(0)
    return output

# ==========================
# INTERFACE STREAMLIT
# ==========================

def app_main():
    st.set_page_config(
        page_title="Catalogue destockage avec images",
        layout="wide"
    )

    st.title("🧠 Catalogue de destockage avec images (textile, chaussures, parfums, etc.)")

    st.header("1. Clé SerpAPI")
    serpapi_key = serpapi_key_selector()

    col1, col2 = st.columns(2)
    with col1:
        max_images = st.slider(
            "Nombre d'images par produit",
            min_value=1,
            max_value=len(IMAGE_COLS),
            value=3,
            step=1,
        )
    with col2:
        sleep_seconds = st.slider(
            "Pause entre chaque requête (secondes, hors mode test)",
            min_value=0.0,
            max_value=1.0,
            value=0.2,
            step=0.1,
            help="Permet de lisser les appels à l'API si tu as beaucoup de lignes.",
        )

    st.header("2. Mode test (pour ajuster les paramètres)")

    test_mode = st.checkbox(
        "Activer le mode test (ne traiter que les X premières lignes et afficher le log détaillé)",
        value=True,
    )

    max_rows_test = st.number_input(
        "Nombre de lignes à traiter en mode test",
        min_value=1,
        value=50,
        step=1,
    )

    st.header("3. Upload du fichier Excel")

    # Nouveau : nombre de lignes à ignorer en haut du fichier
    rows_to_skip_top = st.number_input(
        "Nombre de lignes à ignorer en haut du fichier (avant la ligne d'en-têtes)",
        min_value=0,
        value=0,
        step=1,
        help=(
            "Exemple : si vos en-têtes commencent à la ligne 3 d'Excel, "
            "indiquez 2 pour ignorer les 2 premières lignes (titres, logos, etc.)."
        ),
    )

    uploaded_file = st.file_uploader(
        "Choisis ton fichier Excel (.xlsx)",
        type=["xlsx"],
    )

    if uploaded_file is None:
        st.info("➡️ Uploade un fichier Excel pour commencer.")
        return

    try:
        # On ignore les X premières lignes, puis la première ligne restante devient la ligne d'en-têtes
        df = pd.read_excel(uploaded_file, skiprows=int(rows_to_skip_top))
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}")
        return
    
    st.success("Fichier bien reçu ✅")
    st.subheader("Aperçu des premières lignes")
    st.dataframe(df.head())

    columns = list(df.columns)

    st.header("4. Options de cache")

    use_global_cache = st.checkbox(
        "Activer le cache global (texte de requête -> images) [fichier image_cache.json]",
        value=True,
    )

    use_unique_key_cache = st.checkbox(
        "Activer le cache par clé unique (REF, REF+COLOR, etc.) "
        "et fusionner les lignes partageant cette clé dans la sortie",
        value=True,
    )

    unique_key_columns = []
    if use_unique_key_cache:
        unique_key_columns = st.multiselect(
            "Colonnes composant la clé unique (ex : REF, ou REF + Color) :",
            options=columns,
            help=(
                "Exemple : REF seul, ou [REF, COLOR]. "
                "Pour une même clé unique, l'image sera recherchée une seule fois, "
                "puis les lignes seront fusionnées dans la sortie (une ligne par clé unique)."
            ),
        )

    st.header("5. Colonnes à utiliser pour la requête (ordre de priorité)")

    candidate_keywords = [
        "ean", "gtin", "ref", "code", "sku",
        "nom", "name", "brand", "marque", "description", "taille", "couleur"
    ]
    default_cols = [
        col for col in columns
        if any(kw in col.lower() for kw in candidate_keywords)
    ][:4]

    search_columns = st.multiselect(
        "Colonnes utilisées pour construire la requête (dans l'ordre) :",
        options=columns,
        default=default_cols,
        help="L'ordre compte : ex. [marque, nom, REF, description].",
    )

    st.header("6. Paramètres d’insertion des images dans Excel")

    col3, col4, col5 = st.columns(3)
    with col3:
        jpg_quality = st.slider(
            "Qualité JPG (1 = très compressé, 95 = haute qualité)",
            min_value=30,
            max_value=95,
            value=80,
            step=1,
            help="Compromis entre taille de fichier et qualité visuelle."
        )

    with col4:
        max_width = st.slider(
            "Largeur max des images (px)",
            min_value=40,
            max_value=200,
            value=80,
            step=5,
        )

    with col5:
        max_height = st.slider(
            "Hauteur max des images (px)",
            min_value=40,
            max_value=200,
            value=80,
            step=5,
        )

    st.header("7. Lancer le traitement")

    if not serpapi_key:
        st.info("➡️ Renseigne ou sélectionne une clé SerpAPI pour activer le bouton.")
        return

    if not search_columns:
        st.info("➡️ Sélectionne au moins une colonne de recherche.")
        return

    if st.button("🔍 Enrichir avec des images et les insérer dans l'Excel"):
        with st.spinner("Traitement en cours..."):
            # 1) Enrichissement avec images (et caches)
            df_enriched, log_entries = enrich_df_with_images(
                df.copy(),
                serpapi_key=serpapi_key,
                search_columns=search_columns,
                max_images=max_images,
                sleep_seconds=sleep_seconds,
                test_mode=test_mode,
                max_rows_test=max_rows_test,
                use_global_cache=use_global_cache,
                use_unique_key_cache=use_unique_key_cache,
                unique_key_columns=unique_key_columns,
            )
            # 2) Pas de fusion de lignes : on garde toutes les lignes
            df_output = df_enriched
            # 3) Génération de l'Excel avec images intégrées
            excel_buffer = create_excel_with_embedded_images(
                df_output,
                image_cols=IMAGE_COLS[:max_images],
                jpg_quality=jpg_quality,
                max_width=max_width,
                max_height=max_height,
                unique_key_columns=unique_key_columns if use_unique_key_cache else None,
            )
        st.subheader("Aperçu des données enrichies (après fusion éventuelle)")
        st.dataframe(df_output.head())

        if test_mode:
            st.subheader("🧪 Log du mode test (requêtes, source, clés uniques, images)")
            if log_entries:
                log_df = pd.DataFrame(log_entries)
                st.dataframe(log_df)
            else:
                st.write("Aucun log (peut-être pas de requêtes valides).")

        st.download_button(
            label="⬇️ Télécharger le fichier Excel avec images intégrées",
            data=excel_buffer,
            file_name="produits_avec_images_integrees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def main():
    init_session_state()

    if not st.session_state.authenticated:
        login_form()
    else:
        app_main()


if __name__ == "__main__":
    main()

