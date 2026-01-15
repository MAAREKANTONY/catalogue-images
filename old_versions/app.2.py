import io
import time
import requests
import pandas as pd
import streamlit as st
from PIL import Image
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ==========================
# CONFIG GLOBALE
# ==========================

# Colonnes de sortie pour les URLs d'images
IMAGE_COLS = ["image_url_1", "image_url_2", "image_url_3"]


# ==========================
# FONCTIONS MÉTIER
# ==========================

def build_query(row: pd.Series, search_columns):
    """
    Construit la requête de recherche en respectant l'ordre choisi par l'utilisateur.
    search_columns est une liste ordonnée de noms de colonnes.
    """
    parts = []
    for col in search_columns:
        value = row.get(col)
        if pd.notna(value):
            text = str(value).strip()
            if text:
                parts.append(text)
    return " ".join(parts)


def search_images_bing(query: str, api_key: str, endpoint_base: str, max_results: int = 3):
    """
    Interroge l'API Bing Image Search (via Azure AI Services) et retourne une liste d'URLs d'images.
    endpoint_base est ton endpoint du type :
      https://vente-en-lots.cognitiveservices.azure.com
    """
    endpoint = f"{endpoint_base.rstrip('/')}/bing/v7.0/images/search"

    headers = {"Ocp-Apim-Subscription-Key": api_key}
    params = {
        "q": query,
        "count": max_results,
        "imageType": "Photo",
        "safeSearch": "Strict",
    }

    response = requests.get(endpoint, headers=headers, params=params, timeout=15)
    response.raise_for_status()
    data = response.json()

    values = data.get("value", [])
    urls = [item.get("contentUrl") for item in values if item.get("contentUrl")]
    return urls[:max_results]


def enrich_df_with_images(
    df: pd.DataFrame,
    api_key: str,
    endpoint_base: str,
    search_columns,
    max_images: int = 3,
    sleep_seconds: float = 0.2,
):
    """
    Enrichit un DataFrame avec des URLs d'images.
    Ajoute les colonnes IMAGE_COLS si elles n'existent pas.
    Utilise les colonnes search_columns dans l'ordre fourni.
    """
    if not search_columns:
        st.error("Aucune colonne sélectionnée pour la recherche d'images.")
        return df

    # S'assurer que les colonnes de sortie existent
    for i, col in enumerate(IMAGE_COLS[:max_images]):
        if col not in df.columns:
            df[col] = None

    progress_bar = st.progress(0)
    status_text = st.empty()
    total = len(df)

    for idx, (row_index, row) in enumerate(df.iterrows(), start=1):
        query = build_query(row, search_columns)

        if not query.strip():
            status_text.text(f"Ligne {row_index}: requête vide, ignorée.")
        else:
            try:
                urls = search_images_bing(
                    query=query,
                    api_key=api_key,
                    endpoint_base=endpoint_base,
                    max_results=max_images,
                )
                for i, url in enumerate(urls):
                    df.at[row_index, IMAGE_COLS[i]] = url

                status_text.text(
                    f"Ligne {row_index}: {len(urls)} image(s) trouvée(s) pour requête '{query[:60]}...'"
                )
            except Exception as e:
                status_text.text(f"Ligne {row_index}: ERREUR ({e})")

            time.sleep(sleep_seconds)

        progress_bar.progress(idx / total)

    status_text.text("Recherche d'images terminée ✅")
    return df


def create_excel_with_embedded_images(
    df: pd.DataFrame,
    image_cols,
    jpg_quality: int = 80,
    image_width: int = 80,
    image_height: int = 80,
):
    """
    Crée un fichier Excel en mémoire avec :
    - les données du DataFrame,
    - les images téléchargées, converties en JPG,
      insérées dans les cellules correspondant aux colonnes image_cols.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Écrire les données (on garde les URLs pour debug)
        df.to_excel(writer, index=False, sheet_name="Produits")
        wb = writer.book
        ws = wb["Produits"]

        for row_idx, row in df.iterrows():
            # +2 = ligne 1 = header, ligne 2 = première ligne de données
            excel_row = row_idx + 2

            for img_col in image_cols:
                if img_col not in df.columns:
                    continue

                url = row.get(img_col)
                if not isinstance(url, str) or not url.strip():
                    continue

                try:
                    resp = requests.get(url, timeout=15)
                    resp.raise_for_status()

                    pil_img = Image.open(io.BytesIO(resp.content))
                    if pil_img.mode != "RGB":
                        pil_img = pil_img.convert("RGB")

                    img_bytes = io.BytesIO()
                    pil_img.save(
                        img_bytes,
                        format="JPEG",
                        quality=jpg_quality,
                        optimize=True,
                    )
                    img_bytes.seek(0)

                    xl_img = XLImage(img_bytes)
                    xl_img.width = image_width
                    xl_img.height = image_height

                    col_idx = df.columns.get_loc(img_col) + 1  # 1-based index
                    col_letter = get_column_letter(col_idx)
                    cell_ref = f"{col_letter}{excel_row}"

                    ws.add_image(xl_img, cell_ref)

                except Exception:
                    # On ignore l'erreur pour cette image, on passe à la suite
                    continue

    output.seek(0)
    return output


# ==========================
# INTERFACE STREAMLIT
# ==========================

def main():
    st.set_page_config(page_title="Catalogue produit avec images", layout="wide")

    st.title("🧠 Catalogue produit enrichi avec images (intégrées dans Excel)")
    st.write(
        """
        Cet outil permet :
        - d'uploader un fichier Excel produit,
        - de choisir **l'ordre des colonnes** utilisées pour la requête d'image,
        - d'interroger Bing Image Search (Azure AI Services),
        - d'enrichir le fichier avec des URLs,
        - et d'**insérer directement les images JPG compressées dans l'Excel**.
        """
    )

    st.header("1. Configuration Azure / Bing Image Search")

    endpoint_base = st.text_input(
        "Endpoint Azure (sans /bing/v7.0...)",
        value="https://vente-en-lots.cognitiveservices.azure.com",
        help="Copie ici le 'Point de terminaison' affiché dans Azure AI Services."
    )

    api_key = st.text_input(
        "Clé API (CLÉ 1 ou CLÉ 2 Azure AI Services)",
        type="password",
        help="Copie ici la CLÉ 1 (ou 2) de ta ressource Azure AI Services."
    )

    max_images = st.slider(
        "Nombre d'images par produit",
        min_value=1,
        max_value=len(IMAGE_COLS),
        value=3,
        step=1,
    )

    sleep_seconds = st.slider(
        "Pause entre chaque requête (secondes)",
        min_value=0.0,
        max_value=1.0,
        value=0.2,
        step=0.1,
        help="Permet d'éviter de saturer l'API / dépasser les quotas."
    )

    st.header("2. Upload du fichier Excel")

    uploaded_file = st.file_uploader(
        "Choisis ton fichier Excel (.xlsx)",
        type=["xlsx"],
    )

    if uploaded_file is None:
        st.info("➡️ Uploade un fichier Excel pour commencer.")
        return

    # Lecture du fichier
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}")
        return

    st.success("Fichier bien reçu ✅")
    st.subheader("Aperçu des premières lignes")
    st.dataframe(df.head())

    st.header("3. Choix des colonnes de recherche (ordre de priorité)")

    columns = list(df.columns)

    # Heuristique pour pré-sélectionner quelques colonnes probables
    candidate_keywords = [
        "ean", "gtin", "ref", "code", "sku",
        "nom", "name", "brand", "marque", "description"
    ]
    default_cols = [
        col for col in columns
        if any(kw in col.lower() for kw in candidate_keywords)
    ][:4]

    search_columns = st.multiselect(
        "Sélectionne les colonnes à utiliser pour construire la requête (dans l'ordre) :",
        options=columns,
        default=default_cols,
        help=(
            "L'ordre compte : par exemple [marque, nom, EAN, description]. "
            "Ne sélectionne pas les colonnes inutiles (ex : taille si ça ne t'aide pas)."
        ),
    )

    st.header("4. Paramètres d'insertion des images dans Excel")

    jpg_quality = st.slider(
        "Qualité JPG (1 = très compressé, 95 = haute qualité)",
        min_value=30,
        max_value=95,
        value=80,
        step=1,
        help="Compromis entre taille de fichier et qualité visuelle."
    )

    image_width = st.slider(
        "Largeur des images dans Excel (pixels approximatifs)",
        min_value=40,
        max_value=200,
        value=80,
        step=5,
    )

    image_height = st.slider(
        "Hauteur des images dans Excel (pixels approximatifs)",
        min_value=40,
        max_value=200,
        value=80,
        step=5,
    )

    st.header("5. Lancer le traitement")

    if not api_key or not endpoint_base.strip():
        st.info("➡️ Renseigne l'endpoint Azure et la clé API pour activer le bouton.")
        return

    if not search_columns:
        st.info("➡️ Sélectionne au moins une colonne de recherche.")
        return

    if st.button("🔍 Enrichir avec des images et les insérer dans l'Excel"):
        with st.spinner("Traitement en cours..."):
            # 1) Recherche d'images (URLs)
            df_enriched = enrich_df_with_images(
                df.copy(),
                api_key=api_key,
                endpoint_base=endpoint_base,
                search_columns=search_columns,
                max_images=max_images,
                sleep_seconds=sleep_seconds,
            )

            # 2) Création du fichier Excel avec images intégrées
            excel_buffer = create_excel_with_embedded_images(
                df_enriched,
                image_cols=IMAGE_COLS[:max_images],
                jpg_quality=jpg_quality,
                image_width=image_width,
                image_height=image_height,
            )

        st.subheader("Aperçu des données enrichies (URLs d'images)")
        st.dataframe(df_enriched.head())

        st.download_button(
            label="⬇️ Télécharger le fichier Excel avec images intégrées",
            data=excel_buffer,
            file_name="produits_avec_images_integrees.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()

