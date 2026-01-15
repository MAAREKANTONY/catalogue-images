import io
import time
import zipfile
import requests
import pandas as pd
import streamlit as st
from PIL import Image

# ==========================
# CONFIG IMAGES
# ==========================

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


def search_images_bing(query: str, api_key: str, max_results: int = 3):
    """
    Interroge l'API Bing Image Search et retourne une liste d'URLs d'images.
    """
    endpoint = "https://api.bing.microsoft.com/v7.0/images/search"
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
                urls = search_images_bing(query, api_key, max_results=max_images)
                for i, url in enumerate(urls):
                    df.at[row_index, IMAGE_COLS[i]] = url

                status_text.text(
                    f"Ligne {row_index}: {len(urls)} image(s) trouvée(s) pour requête '{query[:60]}...'"
                )
            except Exception as e:
                status_text.text(f"Ligne {row_index}: ERREUR ({e})")

            time.sleep(sleep_seconds)

        progress_bar.progress(idx / total)

    status_text.text("Traitement terminé ✅")
    return df


def download_and_convert_images_to_zip(
    df: pd.DataFrame,
    image_cols,
    jpg_quality: int = 80,
    prefix: str = "img",
):
    """
    Télécharge les images référencées dans image_cols,
    les convertit en JPG compressé, les stocke dans un zip en mémoire,
    et ajoute dans le DataFrame des colonnes <col>_jpg avec le nom du fichier.
    """
    zip_buffer = io.BytesIO()

    # Créer les colonnes pour les noms de fichiers JPG
    for col in image_cols:
        local_col = f"{col}_jpg"
        if local_col not in df.columns:
            df[local_col] = None

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for row_index, row in df.iterrows():
            for col in image_cols:
                url = row.get(col)
                if not isinstance(url, str) or not url.strip():
                    continue

                filename = f"{prefix}_row{row_index}_{col}.jpg"

                try:
                    resp = requests.get(url, timeout=15)
                    resp.raise_for_status()
                    img = Image.open(io.BytesIO(resp.content))

                    if img.mode != "RGB":
                        img = img.convert("RGB")

                    img_bytes = io.BytesIO()
                    img.save(
                        img_bytes,
                        format="JPEG",
                        quality=jpg_quality,
                        optimize=True,
                    )
                    img_bytes.seek(0)

                    zipf.writestr(filename, img_bytes.read())
                    df.at[row_index, f"{col}_jpg"] = filename

                except Exception:
                    # On ignore les erreurs de téléchargement / conversion pour garder le process robuste
                    continue

    zip_buffer.seek(0)
    return df, zip_buffer


# ==========================
# INTERFACE STREAMLIT
# ==========================

def main():
    st.set_page_config(page_title="Catalogue produit avec images", layout="wide")

    st.title("🧠 Catalogue produit enrichi avec images")
    st.write(
        """
        Cet outil permet :
        - d'uploader un fichier Excel produit,
        - de choisir **l'ordre des colonnes** utilisées pour les requêtes,
        - de récupérer des images via une API,
        - d'enrichir l'Excel avec les URLs,
        - et (optionnel) de **convertir les images en JPG compressé** dans un ZIP.
        """
    )

    st.header("1. Configuration API")

    api_key = st.text_input(
        "Clé API Bing Image Search (Azure Cognitive Services)",
        type="password",
        help="Tu peux créer une clé sur le portail Azure, service 'Bing Search v7'.",
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
        help="Permet d'éviter de saturer l'API / dépasser les quotas.",
    )

    st.header("2. Upload du fichier Excel")

    uploaded_file = st.file_uploader(
        "Choisis ton fichier Excel (.xlsx)",
        type=["xlsx"],
    )

    if uploaded_file is not None:
        st.success("Fichier bien reçu ✅")
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Erreur de lecture du fichier Excel : {e}")
            return

        st.subheader("Aperçu des premières lignes")
        st.dataframe(df.head())

        st.header("3. Choix des colonnes de recherche (ordre de priorité)")

        columns = list(df.columns)

        # Heuristique : pré-sélectionner quelques colonnes probables
        candidate_keywords = ["ean", "gtin", "ref", "code", "nom", "name", "brand", "marque", "description"]
        default_cols = [
            col
            for col in columns
            if any(kw in col.lower() for kw in candidate_keywords)
        ][:4]  # on limite le nombre de défauts

        search_columns = st.multiselect(
            "Sélectionne les colonnes à utiliser pour construire la requête (dans l'ordre) :",
            options=columns,
            default=default_cols,
            help=(
                "L'ordre compte : par exemple [Brand, Nom, EAN, Description]. "
                "Ne sélectionne pas les colonnes qui n'aident pas (ex : taille, couleur si non pertinentes)."
            ),
        )

        st.header("4. Options de conversion des images (JPG)")

        convert_images = st.checkbox(
            "Télécharger et convertir les images en JPG compressé",
            value=False,
            help="Si coché, l'outil télécharge les images trouvées, les convertit en JPG et te fournit un ZIP.",
        )

        jpg_quality = st.slider(
            "Qualité JPG (1 = très compressé, 95 = haute qualité)",
            min_value=30,
            max_value=95,
            value=80,
            step=1,
            help="Compromis entre taille de fichier et qualité visuelle.",
        )

        st.header("5. Lancer le traitement")

        if not api_key:
            st.info("➡️ Entre ta clé API pour activer le bouton.")
        elif not search_columns:
            st.info("➡️ Sélectionne au moins une colonne de recherche.")
        else:
            if st.button("🔍 Enrichir avec des images"):
                with st.spinner("Traitement en cours..."):
                    df_enriched = enrich_df_with_images(
                        df.copy(),
                        api_key=api_key,
                        search_columns=search_columns,
                        max_images=max_images,
                        sleep_seconds=sleep_seconds,
                    )

                    zip_buffer = None
                    if convert_images:
                        df_enriched, zip_buffer = download_and_convert_images_to_zip(
                            df_enriched,
                            image_cols=IMAGE_COLS[:max_images],
                            jpg_quality=jpg_quality,
                            prefix="product",
                        )

                st.subheader("Aperçu des données enrichies")
                st.dataframe(df_enriched.head())

                # Préparation du fichier Excel pour téléchargement
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    df_enriched.to_excel(writer, index=False)
                excel_buffer.seek(0)

                st.download_button(
                    label="⬇️ Télécharger le fichier Excel enrichi",
                    data=excel_buffer,
                    file_name="produits_avec_images.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if convert_images and zip_buffer is not None:
                    st.download_button(
                        label="⬇️ Télécharger le ZIP des images JPG",
                        data=zip_buffer,
                        file_name="images_produits_jpg.zip",
                        mime="application/zip",
                    )

    else:
        st.info("➡️ Uploade un fichier Excel pour commencer.")


if __name__ == "__main__":
    main()

