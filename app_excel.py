import streamlit as st
import pandas as pd
import requests
import re
from io import BytesIO
import io
import json
import time


st.set_page_config(page_title="Boite à outils RDB",  page_icon="🎵", layout="centered")
col1, col2 = st.columns([1, 2])
with col1:
    st.image("logo rdb.jpeg", width=160)  

with col2:
    st.title("Boite à outils RDB")

# --- En-tête
st.subheader("💿 Récupérer la Discographie")
st.markdown("Saisir le nom d'un artiste pour exporter toute sa discographie (albums, singles, collaborations) au format Excel.")

# --- API Deezer (gratuite, sans authentification)
DEEZER_API_BASE = "https://api.deezer.com"


def deezer_get(endpoint, params=None):
    """Effectue un appel GET à l'API Deezer avec gestion du rate limiting."""
    url = f"{DEEZER_API_BASE}/{endpoint}"
    r = requests.get(url, params=params)
    data = r.json()
    # Gestion du rate limiting Deezer
    if "error" in data and data["error"].get("code") == 4:
        time.sleep(1)
        r = requests.get(url, params=params)
        data = r.json()
    return data


def deezer_get_all(endpoint, params=None):
    """Récupère toutes les pages d'un endpoint Deezer paginé."""
    if params is None:
        params = {}
    params["limit"] = 100
    params["index"] = 0
    all_items = []
    while True:
        data = deezer_get(endpoint, params)
        items = data.get("data", [])
        if not items:
            break
        all_items.extend(items)
        if "next" not in data:
            break
        params["index"] += len(items)
    return all_items


# --- Fonctions utilitaires
def seconds_to_hhmmss(seconds):
    if seconds is None:
        return ""
    total_seconds = int(seconds)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    secs = total_seconds % 60
    if hours > 0:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    else:
        return f"0:{minutes:02d}:{secs:02d}"

def artist_data_is_main(artist_id, contributors):
    """Vérifie si l'artiste est le premier contributeur (artiste principal)."""
    return contributors and contributors[0]["id"] == artist_id

def get_artist_discography_export(artist_name=None, artist_id=None):
    """
    Récupère toute la discographie d'un artiste (incluant collaborations)
    et retourne un DataFrame au format demandé.
    """

    # --- Si on a un artist_id, on l'utilise directement
    if artist_id:
        try:
            artist = deezer_get(f"artist/{artist_id}")
            if "error" in artist:
                print(f"❌ ID Deezer invalide ou non trouvé : {artist_id}")
                return None
        except Exception as e:
            print(f"❌ ID Deezer invalide ou non trouvé : {artist_id}")
            return None
    else:
        results = deezer_get("search/artist", params={"q": artist_name, "limit": 1})
        items = results.get("data", [])
        if not items:
            print(f"❌ Aucun artiste trouvé pour '{artist_name}'.")
            return None
        artist = items[0]
        artist_id = artist["id"]

    print(f"🎵 Artiste trouvé : {artist['name']} (Deezer ID: {artist_id})")

    # --- Récupération de tous les albums de cet artiste
    albums = deezer_get_all(f"artist/{artist_id}/albums")

    # --- Suppression des doublons
    seen = set()
    unique_albums = []
    for a in albums:
        if a["id"] not in seen:
            seen.add(a["id"])
            unique_albums.append(a)

    print(f"💿 {len(unique_albums)} albums/singles trouvés (incluant collaborations).")

    # --- Boucle sur tous les albums
    all_rows = []
    for i, album in enumerate(unique_albums, start=1):
        print(f"→ Traitement {i}/{len(unique_albums)} : {album.get('title', '')[:50]}")

        album_id = album["id"]
        album_info = deezer_get(f"album/{album_id}")
        if "error" in album_info:
            continue
        upc = str(album_info.get("upc", ""))
        release_date = album_info.get("release_date", "")
        album_type = album_info.get("record_type", "")
        label = album_info.get("label", "")
        album_url = album_info.get("link", "")

        tracks_data = deezer_get_all(f"album/{album_id}/tracks")
        for track in tracks_data:
            # Récupérer les contributeurs du track
            contributors = track.get("contributors", [])
            # Si pas de contributors dans le listing, on récupère le détail du track
            if not contributors:
                track_detail = deezer_get(f"track/{track['id']}")
                if "error" in track_detail:
                    continue
                contributors = track_detail.get("contributors", [])
            
            artist_ids = [c["id"] for c in contributors]
            artist_names = [c["name"] for c in contributors]

            # Filtre strict : seulement les morceaux où l'artiste ID correspond exactement
            if artist_id not in artist_ids:
                continue

            # Récupérer le détail du track pour ISRC et durée complète
            track_data = deezer_get(f"track/{track['id']}")
            if "error" in track_data:
                continue
            duration = seconds_to_hhmmss(track_data.get("duration"))
            isrc = track_data.get("isrc", "")
            track_url = track_data.get("link", "")
            role = "Main artist" if artist_data_is_main(artist_id, contributors) else "Featured artist"

            all_rows.append({
                'ARTIST NAME': "",
                'ALIAS': artist["name"],  
                'RELEASE ARTIST / GROUP': ", ".join(artist_names),
                'ALBUM TITLE': album_info.get("title", ""),
                'TRACK TITLE': track_data.get("title", ""),
                'Version': "",
                'ISRC CODE': isrc,
                'UPC': upc,
                'Duration': duration,
                'LABEL NAME': label,
                'LABEL COUNTRY': "",
                'YEAR OF RECORDING': release_date[:4] if release_date else "",
                'COUNTRY OF RECORDING': "",
                'RELEASE FORMAT': "",
                'RELEASE TYPE': album_type.capitalize() if album_type else "",
                'ROLE': role,
                'INSTRUMENT(S) / VOCALS': "",
                'PROOF (URL link)': track_url or album_url,
            })

    columns = ['ARTIST NAME', 'ALIAS', 'RELEASE ARTIST / GROUP', 'ALBUM TITLE', 
               'TRACK TITLE', 'Version', 'ISRC CODE', 'Duration', 'LABEL NAME', 
               'LABEL COUNTRY', 'YEAR OF RECORDING', 'COUNTRY OF RECORDING', 
               'RELEASE FORMAT', 'RELEASE TYPE', 'ROLE', 'INSTRUMENT(S) / VOCALS', 
               'PROOF (URL link)', 'UPC','ROYALTY SHARE']

    df = pd.DataFrame(all_rows, columns=columns)
    print(f"✅ {len(df)} morceaux trouvés pour {artist['name']}.")
    return df

# --- Interface utilisateur
artist_name = st.text_input("Nom de l'artiste :")
artist_id_input = st.text_input("Ou ID Deezer de l'artiste :", 
                                placeholder="Ex : 27")

if st.button("🎶 Rechercher et générer"):
    if not artist_name.strip() and not artist_id_input.strip():
        st.warning("Merci de saisir un nom ou un ID Deezer d'artiste.")
    else:
        with st.spinner("Recherche en cours sur Deezer..."):
            if artist_id_input.strip():
                # 🔹 Si l'utilisateur donne un ID Deezer, on l'utilise directement
                try:
                    artist = deezer_get(f"artist/{artist_id_input.strip()}")
                    if "error" in artist:
                        st.error("❌ ID Deezer invalide ou non trouvé.")
                        st.stop()
                    artist_name = artist["name"]
                    st.info(f"🎵 Artiste trouvé via ID : **{artist_name}**")
                    df = get_artist_discography_export(artist_id=int(artist_id_input.strip()))
                except Exception:
                    st.error("❌ ID Deezer invalide ou non trouvé.")
                    st.stop()
            else:
                # 🔹 Sinon, recherche classique par nom
                df = get_artist_discography_export(artist_name=artist_name.strip())

        if df is not None:
            st.success(f"✅ {len(df)} morceaux trouvés pour {artist_name} !")

            st.dataframe(df.head(10))

            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                label="📥 Télécharger en Excel",
                data=output,
                file_name=f"{artist_name}_discography.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Suite  

st.subheader("🎡​Moulinette Droits Voisins")
st.markdown("""
Importer le répertoire source pour le convertir aux formats requis par:  
**Spedidam**, **Playright**, **SwissPerf**, **SENA**, **AIE**, **Artisti** et **EJI**.
""")

# Upload du fichier Excel
uploaded_file = st.file_uploader("Choisissez un fichier Excel source", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Lecture du fichier Excel
        df_source = pd.read_excel(uploaded_file, header=0)
        df_source["YEAR OF RECORDING"] = df_source["YEAR OF RECORDING"].astype("Int64").astype(str)
        
        # 👉 Sélection des onglets à inclure
        options = ["SPED", "Playright", "SwissPerf", "SENA", "AIE", "Artisti", "Artisti_2", "EJI"]
        selected_tabs = st.multiselect(
            "✅ Choisissez les onglets à inclure dans l'export :",
            options,
            default=options,
            key="export_tabs"  # clé unique
        )

        #SPED

        # Fonction pour découper la durée
        def split_duration(duration):
            try:
                h, m, s = duration.split(":")
                return int(h), int(m), int(s)
            except:
                return None, None, None

        # Construire df_sped
        df_sped = pd.DataFrame()

        # Colonnes obligatoires
        df_sped["Artiste représenté\nRepresented artist"] = df_source['ARTIST NAME']

        df_sped['Pseudonyme\nStage Name'] = df_source["ALIAS"]

        # Premier instrument = premier mot des instruments
        df_sped["Premier instrument\nFirst Function on the title"] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: str(x).split(",")[0] if pd.notnull(x) else ""
        )

        # Autres instruments = tout sauf le premier (après la 1ère virgule)
        df_sped["Autres instruments\nOther  Functions on the title"] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: ",".join([i.strip() for i in str(x).split(",")[1:]]) if pd.notnull(x) and "," in str(x) else ""
        )

        # 1️⃣ Statut Artiste Principal
        def is_main_artist_partial(row):
            main = str(row.get("RELEASE ARTIST / GROUP", "")).strip().lower()
            pseudos = str(row.get("ALIAS", "")).strip().lower().split(";")
            pseudos = [p.strip() for p in pseudos if p.strip()]
            rep = str(row.get("ARTIST NAME", "")).strip().lower()
            role = str(row.get("ROLE", "")).strip().upper()

            # Si ROLE = FA → artiste principal
            if role == "FA":
                return "yes"

            # Vérifie si au moins un pseudo ou le nom de l'artiste est dans le nom du groupe
            for p in pseudos + [rep]:
                if p and p in main:
                    return "yes"
            return "no"


        # Appliquer la fonction
        df_sped["Statut Artiste Principal\nMain artist\n(yes or no)"] = df_source.apply(is_main_artist_partial, axis=1)


        # 2️⃣ Nom de l'artiste principal : toujours le nom du groupe
        df_sped["Nom de l'artiste principal\nName of Main Artist"] = df_source["RELEASE ARTIST / GROUP"]


        # 3️⃣ Nom du Groupe principal : seulement si ce n'est PAS un artiste principal
        df_sped["Nom du Groupe principal\nName of Main Group"] = df_source["RELEASE ARTIST / GROUP"].where(
            df_sped["Statut Artiste Principal\nMain artist\n(yes or no)"] == "no", ""
        )


        # 4️⃣ Membre du groupe : si ROLE = FA et que (ALIAS ou ARTIST NAME) ≠ RELEASE ARTIST / GROUP
        def map_role_swissperf(row):
            role = str(row.get("ROLE", "")).strip().upper()  # FA ou NFA
            alias_field = str(row.get("ALIAS", "")).strip().lower()
            artist = str(row.get("ARTIST NAME", "")).strip().lower()
            group = str(row.get("RELEASE ARTIST / GROUP", "")).strip().lower()

            # Si ROLE n'est pas valide
            if role not in ["FA", "NFA"]:
                return ""

            # Nettoyage du nom du groupe pour faciliter la détection
            # On supprime les mots typiques de featuring
            group_clean = re.sub(r'\b(feat\.?|ft\.?|avec|and|&|,|;)\b', '/', group)
            group_parts = [g.strip() for g in group_clean.split("/") if g.strip()]

            # Prépare la liste de comparaison (artiste + alias)
            aliases = [a.strip() for a in alias_field.split(";") if a.strip()]
            all_names = [artist] + aliases

            # 🔹 Si un des noms (artist ou alias) correspond à un des noms principaux → soliste
            for name in all_names:
                for g in group_parts:
                    if g.startswith(name) or name in g:
                        return f"{role}/S"

            # 🔹 Sinon → membre d'une formation
            return f"{role}/MF"

        df_sped["Membre du groupe\nGroup member"] = df_source.apply(map_role_swissperf, axis=1)
        # Release Title = Release type
        df_sped["Titre général\nRelease Title"] = df_source["ALBUM TITLE"]

        # Titre de l'enregistrement
        df_sped["Titre de l'enregistrement\nTrack Title"] = df_source["TRACK TITLE"]

        # Durée en Heure, Minute, Seconde
        # Création des colonnes de durée pour df_sped uniquement
        heures, minutes, secondes, durees = [], [], [], []

        for duree in df_source['Duration']:
            d = str(duree).strip()

            # On ne modifie PAS df_source, juste la version pour df_sped
            if ":" in d:
                parts = d.split(":")
                if len(parts) == 3:  # Déjà HH:MM:SS
                    h, m, s = parts
                elif len(parts) == 2:  # Format MM:SS ou M:SS
                    h = "00"
                    m = parts[0].zfill(2)
                    s = parts[1].zfill(2)
                else:
                    h, m, s = "00", "00", "00"
            else:
                h, m, s = "00", "00", "00"

            heures.append(h)
            minutes.append(m)
            secondes.append(s)
            durees.append(f"{h}:{m}:{s}")

        # Appliquer uniquement sur df_sped
        df_sped["Durée\nDuration"] = durees
        df_sped["Heure\nHour"] = heures
        df_sped["Minute\nMinute"] = minutes
        df_sped["Seconde\nSecond"] = secondes

        # Producteur Original
        df_sped["Producteur Original\nOriginal Record Label"] = df_source['LABEL NAME']

        # Nationalité du producteur d'origine
        df_sped["Nationalité du producteur d'origine\nNationality of Producer"] = df_source['LABEL COUNTRY']

        # Pays d'enregistrement
        df_sped["Pays d'enregistrement\nCountry of Recording"] = df_source['COUNTRY OF RECORDING']

        # Date d'enregistrement (vide par défaut car on n'a que l'année)
        df_sped["date d'enregistrement\ndate of recording\n(dd/mm/yyyy)"] = df_source['YEAR OF RECORDING']

        # Année de première publication
        df_sped["Année de première publication\nYear of first publication\n(dd/mm/yyyy)"] = df_source['YEAR OF RECORDING']

        #ISRC 
        df_sped["ISRC"] = df_source['ISRC CODE']
            
        colonnes_sped = [
            'Artiste représenté\nRepresented artist',
            'Pseudonyme\nStage Name',
            'Premier instrument\nFirst Function on the title',
            'Autres instruments\nOther  Functions on the title',
            'Statut Artiste Principal\nMain artist\n(yes or no)',
            'Membre du groupe\nGroup member',
            'Nom de l\'artiste principal\nName of Main Artist',
            'Nom du Groupe principal\nName of Main Group',
            'Nom de l\'Orchestre/Chœur\nName of Orchester/Choir',
            'Support de diffusion\nType of Media',
            'Titre général\nRelease Title',
            'Titre de l\'enregistrement\nTrack Title',
            'Compositeur\nComposer\n(only for Classical music)',
            'Durée\nDuration',
            'Heure\nHour',
            'Minute\nMinute',
            'Seconde\nSecond',
            'Producteur Original\nOriginal Record Label',
            'Nationalité du producteur d\'origine\nNationality of Producer',
            'ISRC',
            'Studio ou lieu d\'enregistrement\nStudio or place of recording',
            'Pays d\'enregistrement\nCountry of Recording',
            'date d\'enregistrement\ndate of recording\n(dd/mm/yyyy)',
            'Année de première publication\nYear of first publication\n(dd/mm/yyyy)',
            'Année de réédition \nYear of reissue\n(dd/mm/yyyy)',
            'Style de musique\nMusical style'
        ]

        # Ajouter colonnes manquantes vides
        for col in colonnes_sped:
            if col not in df_sped.columns:
                df_sped[col] = ""

        # Réordonner
        df_sped = df_sped[colonnes_sped]

        instru_dict = {'Keyboards': ['Accordion',
        'Anglo-Concertina',
        'Bandoneon',
        'Button Accordion',
        'Calliope',
        'Celeste',
        'Chamberlin',
        'Clavichord',
        'Clavier',
        'Clavinet',
        'Electric Piano',
        'Electronium',
        'English Concertina',
        'Fortepiano',
        'Glass Harmonica',
        'Hammond Organ',
        'Harmonium',
        'Harpsichord',
        'Hurdy-Gurdy',
        'Keyboards',
        'Mellotron',
        'Novachord',
        'Organ',
        'Piano',
        'Piano Accordion',
        'Player Piano',
        'Spinet',
        'Stylophone',
        'Synthesizer',
        'Torader',
        'Virginal'],
        'Mechanical': ['Barrel-Organ',
        'Beat Box',
        'Chapman Stick',
        'Lyricon',
        'Music Box',
        'Omnichord',
        'Theremin'],
        'Percussion': ['Achere',
        'Afuche',
        'Agogo',
        'Balafon',
        'Bass Drum',
        'Bells',
        'Berimbau',
        'Bodhran',
        'Bombo Leguero',
        'Bones',
        'Bongos',
        'Bottle',
        'Cabasa',
        'Carillon',
        'Caxixi',
        'Chekere',
        'Chimes',
        'Chinese Drum',
        'Clapping',
        'Claves',
        'Congas',
        'Congos',
        'Cowbell',
        'Cuica',
        'Cymbals',
        'Daf',
        'Daire',
        'Darabuka',
        'Dhol',
        'Dholak',
        'Djembe',
        'Dohol',
        'Donno',
        'Drums/Drumkit',
        'Duggi',
        'Dumbak',
        'Finger Cymbals',
        'Frying-Pan/Frigideira',
        'Ganga',
        'Ganza',
        'Ghatam',
        'Glockenspiel',
        'Gong',
        'Guayo',
        'Guiro',
        'Hoshu',
        'Jug',
        'Kalimba',
        'Kanjira',
        'Kartal',
        'Korintsana',
        'Lamellaphone',
        'Log Drum',
        'Lokole',
        'Lukeme',
        'Luo',
        'Maraccas',
        'Marimba',
        'Marimbula',
        'Mbira',
        'Mrdanga',
        'Pandeiro',
        'Pano',
        'Percussion',
        'Quinto',
        'Rattle',
        'Reco Reco',
        'Repinique',
        'Rnga',
        'Rol-Mo',
        'Sansa',
        'Saw',
        'Scraper',
        'Sekere',
        'Siren',
        'Snaredrum',
        'Spoons',
        'Steel Drum',
        'Surdo',
        'Tabla',
        'Taiko',
        'Tamborim',
        'Tambourine',
        'Tamtam',
        'Tapan',
        'Tarol/Caixa',
        'Tavil',
        'Templeblocks',
        'Timpani',
        'Triangle',
        'Vibraphone',
        'Vibra-Slap',
        'Washboard',
        'Whip',
        'Wind Machine',
        'Wobble Board',
        'Wood Chimes',
        'Woodblock',
        'Xylophone',
        'Zarb',
        'Zeze',
        'Zil',
        'Zirbaghali'],
        'Strings': ['Acoustic Guitar',
        'Andean Harp',
        'Appalachian Dulcimer',
        'Autoharp',
        'Baglama',
        'Balalaika',
        'Bandurria',
        'Banjo',
        'Baryton',
        'Bass Guitar',
        'Bouzouki',
        'Cello',
        'Charango',
        'Cheng',
        'Chinese Harp',
        'Chitarrone',
        'Cimbalon',
        'Citole',
        'Cittern',
        'Crwth',
        'Cuatro',
        'Dan Tranh',
        'Dombra',
        'Double Bass',
        'Dulcimer',
        'Dutar',
        'Electric Guitar',
        'Erhu',
        'Fiddel',
        'Five String Banjo',
        'Gadulka',
        'Gardon',
        'Gimbri',
        'Gittern',
        'Guitar',
        'Guitarra',
        'Guitarron',
        'Hammered Dulcimer',
        'Hardangerfiddle',
        'Harp',
        'Hawaii Guitar',
        'Hoddu',
        'Huapanguera',
        'Imitation Bass',
        'Indian Harp',
        'Irish Harp',
        'Jumbush',
        'Kacapi',
        'Kantele',
        'Kayagum',
        'Kemence',
        'Kemene',
        'Khalam',
        'Kora',
        'Koto',
        'Kra',
        'Langeleik',
        'Lute',
        'Lyre',
        'Mandocello',
        'Mandola',
        'Mandolin',
        'Mandore',
        'Mouth Bow',
        'Njarka',
        'Nkoni',
        'Nyatitit',
        'Nycleharpe',
        'Oud',
        'Pedabro',
        'Psaltery',
        'Qanun',
        'Quinto',
        'Rabab',
        'Rababa',
        'Rebec',
        'Santur',
        'Sarangi',
        'Saz',
        'Sitar',
        'Steel Guitar',
        'Strings',
        'Surbahar',
        'Surmandal',
        'Tambura',
        'Tenor Banjo',
        'Theorbo',
        'Tiompan',
        'Tiple',
        'Tsuguru Samisen',
        'Ud',
        'Ukulele',
        'Valiha',
        'Vena',
        'Vihuela',
        'Vina',
        'Viola',
        'Viola Da Gamba',
        'Viola Damore',
        'Violin',
        'Yangqin',
        'Zither'],
        'TunedPercussion': ['Hand Bells', 'Timbales', 'Tubular Bells'],
        'Vocal': ['Alto',
        'Baritone',
        'Bass',
        'Bass-Baritone',
        'Beat Box',
        'Contralto',
        'Counter Tenor',
        'Falsetto',
        'Joik',
        'Mezzo-Soprano',
        'Rap',
        'Soprano',
        'Spoken Word',
        'Tenor',
        'Treble',
        'Vocal',
        'Vocoder',
        'Whistling',
        'Yodel'],
        'Wind': ['Aeolin Harp',
        'Alphorn',
        'Alto Clarinet',
        'Alto Flute',
        'Alto Saxophone',
        'Andean Flute',
        'Antara',
        'Arghul',
        'Bagpipes',
        'Bansuri',
        'Baritone',
        'Baritone Saxophone',
        'Bass Clarinet',
        'Bass Flute',
        'Bass Saxophone',
        'Bass Trombone',
        'Basset Horn',
        'Bassoon',
        'Border Pipes',
        'Bottle',
        'Bugle',
        'Bukkehorn',
        'C Melody Saxophone',
        'Chinese Flute',
        'Chiwata',
        'Clarinet',
        'Comb',
        'Contrabass Clarinet',
        'Cor Anglais',
        'Cornet',
        'Cornett',
        'Crumhorn',
        'Curtal',
        'Didjeridu',
        'Dobro',
        'Double Fluite',
        'Duduk',
        'Dung-Chen',
        'Electronic Wind Instr',
        'English Bagpipes',
        'English Horn',
        'Euphonium',
        'Fanfare Trumpet',
        'Fife',
        'Flageolet',
        'Floyera',
        'Flugelhorn',
        'Flute',
        'French Bagpipes',
        'French Horn',
        'Fujara',
        'Gajde',
        'Galo',
        'Genggong',
        'Harmonica',
        'Helico',
        'Highland Pipes',
        'Horn',
        'Jagd Horn',
        'Japanese Wooden Flute',
        'Jazz Horn',
        'Jews Harp',
        'Jug',
        'Kaval',
        'Kazoo',
        'Kena',
        'Lakoto Flute',
        'Lur',
        'Mellophone',
        'Melodeon',
        'Melodica',
        'Mizmar',
        'Mouth Organ',
        'Munnharpe',
        'Musette',
        'Nagasvaram',
        'Nanga',
        'Nay',
        'Northumbrian Small Pipes',
        'Oboe',
        'Oboe Damore',
        'Ocarina',
        'Ophicleide',
        'Organ',
        'Panpipes',
        'Piccolo',
        'Pulankulai',
        'Quitls',
        'Racket',
        'Rauschpfeife',
        'Recorder',
        'Rondadores',
        'Sackbut',
        'Saxello',
        'Saxophone',
        'Seljeflyte',
        'Serpent',
        'Shakuhachi',
        'Shanai',
        'Shavvm',
        'Sheng',
        'Sopranino',
        'Soprano Saxophone',
        'Sordone',
        'Sousaphone',
        'Swannee Whistle',
        'Tarogato',
        'Tenor Horn',
        'Tenor Saxophone',
        'Trombone',
        'Trumpet',
        'Trumpet Marine',
        'Tuba',
        'Tusseflute',
        'Uilleann Pipes',
        'Valve Trombone',
        'Waldhorn',
        'Whistle',
        'Wooden Flute',
        'Zampona',
        'Zurla',
        'Zurna']}

        # Construire df_playright
        df_playright = pd.DataFrame()

        # Construire df_playright
        df_playright = pd.DataFrame()

        # Colonnes obligatoires
        df_playright["Results"] = ""  # vide par défaut
        df_playright["PerformerName"] = df_source.get('ARTIST NAME', "")
        df_playright["AffiliationNumber"] = ""
        df_playright["AlbumTitle"] = df_source.get('ALBUM TITLE', "") 
        df_playright["TitleTrack"] = df_source.get('TRACK TITLE', "")
        df_playright["MainArtist"] = df_source.get('RELEASE ARTIST / GROUP', "")
        df_playright["RecordingYear"] = df_source.get('YEAR OF RECORDING', "")
        df_playright["ProductionName"] = df_source.get('LABEL NAME', "")
        df_playright["CountryOfProduction"] = df_source.get('LABEL COUNTRY', "")
        df_playright["CountryOfMastering"] = df_source.get('COUNTRY OF RECORDING', "")

        # PerformerRole : main artist / featured artist (logique partielle)
        def performer_role_partial(row):
            main = str(row.get('RELEASE ARTIST / GROUP', "")).strip().lower()
            pseudos = str(row.get('ALIAS', "")).strip().lower().split(";")
            pseudos = [p.strip() for p in pseudos if p.strip()]
            rep = str(row.get('ARTIST NAME', "")).strip().lower()
            for p in pseudos + [rep]:
                if p and p in main:
                    return "Main Artist"
            return "Featured Artist"

        # Mapping ROLE -> PerformerRole
        role_map = {
            "NFA": "C",  # musicien de session
            "FA": "A"    # contrat d'artiste
        }

        df_playright["PerformerRole"] = df_source['ROLE'].apply(
            lambda x: role_map.get(str(x).strip(), "")
        )

        df_playright["PerformerRole2"] = ""  # vide par défaut
        import unicodedata

        # Normalisation du texte
        def normalize_text(text):
            if not text:
                return ""
            text = str(text).strip().lower()
            text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
            return text

        # Trouver la famille de l'instrument dans le dictionnaire
        def find_instrument_family(instr):
            instr_norm = normalize_text(instr)
            if not instr_norm:
                return ""
            for family, instruments in instru_dict.items():
                for i in instruments:
                    i_norm = normalize_text(i)
                    # correspondance simple
                    if instr_norm == i_norm or instr_norm in i_norm or i_norm in instr_norm:
                        return family
            return ""  # si non trouvé

        # Remplissage dynamique des colonnes Instruments et InstrumentType jusqu'à 5
        max_instruments = 5

        for idx, row in df_source.iterrows():
            instruments = [x.strip() for x in str(row['INSTRUMENT(S) / VOCALS']).split(",") if x.strip()]
            for i in range(max_instruments):
                col_instr = f"Instruments{i+1}" if i > 0 else "Instruments"
                col_type = f"InstrumentType{i+1}" if i > 0 else "InstrumentType"
                if i < len(instruments):
                    instr = instruments[i]
                    df_playright.at[idx, col_instr] = instr
                    df_playright.at[idx, col_type] = find_instrument_family(instr)
                else:
                    df_playright.at[idx, col_instr] = ""
                    df_playright.at[idx, col_type] = ""

        # Colonnes dans l'ordre final
        colonnes_playright = [
            'Results', 'PerformerName', 'AffiliationNumber', 'AlbumTitle',
            'TitleTrack', 'MainArtist', 'RecordingYear', 'ProductionName',
            'CountryOfProduction', 'CountryOfMastering', 'PerformerRole',
            'PerformerRole2', 'InstrumentType', 'Instruments', 'InstrumentType2',
            'Instruments2', 'InstrumentType3', 'Instruments3', 'InstrumentType4',
            'Instruments4', 'InstrumentType5', 'Instruments5'
        ]

        # Ajouter colonnes manquantes vides
        for col in colonnes_playright:
            if col not in df_playright.columns:
                df_playright[col] = ""


        # Réordonner
        df_playright = df_playright[colonnes_playright]

        # Nouveau DataFrame cible
        df_swissperf = pd.DataFrame()

        # Titre (morceau/chanson/plage) = Track title
        df_swissperf["Titre (morceau/chanson/plage)"] = df_source['TRACK TITLE']

        # Interprète principal/formation = Main artist / group
        df_swissperf["Interprète principal/formation"] = df_source['RELEASE ARTIST / GROUP']

        # Instrument(s) = Instruments / Vocals
        df_swissperf["Instrument(s)"] = df_source['INSTRUMENT(S) / VOCALS']


        def map_role_swissperf(row):
            release_artists = str(row.get('RELEASE ARTIST / GROUP', "")).strip().lower()
            artist_name = str(row.get('ARTIST NAME', "")).strip().lower()
            aliases = [a.strip().lower() for a in str(row.get('ALIAS', "")).split(";") if a.strip()]
            instrument = str(row.get('INSTRUMENT(S) / VOCALS', "")).strip().lower()
            role_prefix = str(row.get('ROLE', "")).upper()  # FA ou NFA

            first_artist = re.split(r'[&,]', release_artists)[0].strip()

            # --- Cas Featured Artist ---
            if role_prefix == "FA":
                if first_artist == artist_name or first_artist in aliases:
                    return "FA/S"
                else:
                    return "FA/MF"

            # --- Cas Non-Featured Artist ---
            elif role_prefix == "NFA":
                if "vocal" in instrument or "voice" in instrument or "chant" in instrument:
                    return "NFA/CS"  # chanteur de studio
                else:
                    return "NFA/MS"  # musicien de studio

            # --- Par défaut ---
            return ""

        # Appliquer sur le DataFrame
        df_swissperf["Rôle(s)\nFA/S, FA/CO, FA/MF\nou\nNFA/MS,  NFA/CS, NFA/MF\net/ou\nAP"] = df_source.apply(map_role_swissperf, axis=1)





        # Pays d'enregistrement
        df_swissperf["Pays d'enregistrement"] = df_source['COUNTRY OF RECORDING']

        # Année d'enre-gistre-ment
        df_swissperf["Année d'enre-gistre-ment"] = df_source['YEAR OF RECORDING']

        # Pays de publication = Label country
        df_swissperf["Pays de publication"] = df_source['LABEL COUNTRY']

        # Année de publica-tion = Year of recording (par défaut)
        df_swissperf["Année de publica-tion"] = df_source['YEAR OF RECORDING']

        # Album = Release type
        df_swissperf["Album"] = df_source['ALBUM TITLE']

        # Label
        df_swissperf["Label"] = df_source['LABEL NAME']

        # Justificatif (Obligation pour AP) = vide
        df_swissperf["Justifi-catif\n(Obliga-tion pour AP)"] = ""

        # Unnamed: 11 = vide
        df_swissperf["Unnamed: 11"] = ""

        # Durée d'enregistrement en min:sec
        def convert_duration(d):
            if pd.isnull(d):
                return ""
            parts = str(d).split(":")
            if len(parts) == 3:  # HH:MM:SS
                h, m, s = parts
                total_min = int(h) * 60 + int(m)
                return f"{total_min}:{s.zfill(2)}"
            elif len(parts) == 2:  # MM:SS
                m, s = parts
                return f"{int(m)}:{s.zfill(2)}"
            return ""

        df_swissperf["Durée d'enregistrement\nen min:sec"] = df_source['Duration'].apply(convert_duration)

        # Réordonner les colonnes
        colonnes_swissperf = [
            'Titre (morceau/chanson/plage)',
            'Interprète principal/formation',
            'Instrument(s)',
            'Rôle(s)\nFA/S, FA/CO, FA/MF\nou\nNFA/MS,  NFA/CS, NFA/MF\net/ou\nAP',
            'Pays d\'enregistrement',
            'Année d\'enre-gistre-ment',
            'Pays de publication',
            'Année de publica-tion',
            'Album',
            'Label',
            'Justifi-catif\n(Obliga-tion pour AP)',
            'Unnamed: 11',
            'Durée d\'enregistrement\nen min:sec'
        ]

        df_swissperf = df_swissperf[colonnes_swissperf]

        # Dictionnaire noms internationaux -> code ISO 3166-1 alpha-2
        pays_codes = {
            "France": "FR",
            "Germany": "DE",
            "Spain": "ES",
            "Italy": "IT",
            "Portugal": "PT",
            "United Kingdom": "GB",
            "Ireland": "IE",
            "Belgium": "BE",
            "Netherlands": "NL",
            "Switzerland": "CH",
            "Austria": "AT",
            "Sweden": "SE",
            "Norway": "NO",
            "Denmark": "DK",
            "Finland": "FI",
            "Poland": "PL",
            "Czechia": "CZ",
            "Hungary": "HU",
            "Romania": "RO",
            "Bulgaria": "BG",
            "Greece": "GR",
            "Russia": "RU",
            "Ukraine": "UA",
            "United States": "US",
            "Canada": "CA",
            "Mexico": "MX",
            "Brazil": "BR",
            "Argentina": "AR",
            "Colombia": "CO",
            "Chile": "CL",
            "Peru": "PE",
            "Venezuela": "VE",
            "Uruguay": "UY",
            "Bolivia": "BO",
            "Paraguay": "PY",
            "Ecuador": "EC",
            "Australia": "AU",
            "New Zealand": "NZ",
            "Japan": "JP",
            "China": "CN",
            "India": "IN",
            "Pakistan": "PK",
            "Bangladesh": "BD",
            "Sri Lanka": "LK",
            "Indonesia": "ID",
            "Malaysia": "MY",
            "Philippines": "PH",
            "Singapore": "SG",
            "Thailand": "TH",
            "Vietnam": "VN",
            "South Korea": "KR",
            "North Korea": "KP",
            "Turkey": "TR",
            "Israel": "IL",
            "Iran": "IR",
            "Iraq": "IQ",
            "Saudi Arabia": "SA",
            "United Arab Emirates": "AE",
            "Egypt": "EG",
            "South Africa": "ZA",
            "Nigeria": "NG",
            "Kenya": "KE",
            "Ghana": "GH",
            "Senegal": "SN",
            "Morocco": "MA",
            "Algeria": "DZ",
            "Tunisia": "TN",
            "Libya": "LY"
        }



        # Fonction simple pour normaliser la casse et les espaces
        def normalize_country_simple(name):
            if pd.isna(name):
                return ""
            return name.strip().lower()

        # Normaliser le dictionnaire
        pays_codes_norm = {k.lower(): v for k, v in pays_codes.items()}


        # Créer le DataFrame de destination
        df_sena = pd.DataFrame()

        # Colonnes obligatoires (*)
        df_sena['SENA_NUMBER*'] = ""  # vide par défaut
        # Colonne ROLE* : MA ou S
        def map_role_sena(row):
            role = str(row.get("ROLE", "")).strip().upper()
            if role == "FA" or role == "Main artist" or role == "Featured artist":
                return "MA"  # Main Artist
            elif role == "NFA":
                return "S"   # Session / Supporting
            else:
                return ""
        df_sena['ROLE*'] = df_source.apply(map_role_sena, axis=1)
        df_sena['ARTIST*'] = df_source['RELEASE ARTIST / GROUP']
        df_sena['TITLE*'] = df_source['TRACK TITLE']
        df_sena['VERSION'] = df_source['Version']
        df_sena['ISRC'] = df_source['ISRC CODE']
        df_sena['TRACKID'] = ""  # vide par défaut

        # Instruments
        # Instrument principal = premier instrument
        df_sena['INSTRUMENTS'] = df_source['INSTRUMENT(S) / VOCALS'].apply(lambda x: str(x).split(",")[0] if pd.notnull(x) else "")
        # Nombre d'instruments
        df_sena['NO_OF_INSTRUMENTS'] = ""

        # Catégorie selon ROLE
        df_sena['CATEGORY*'] = "M"

        # Label et genre
        df_sena['LABEL'] = df_source['LABEL NAME']
        df_sena['GENRE*'] = "P"  # à remplir si disponible
        df_sena['COMPOSER'] = ""  # vide par défaut

        # Durée (HH:MM:SS → convertir en secondes ou garder textuel)
        def convert_duration(duration):
            d = str(duration).strip()
            if ":" in d:
                parts = d.split(":")
                if len(parts) == 3:
                    h, m, s = parts
                elif len(parts) == 2:
                    h = "0"
                    m, s = parts
                else:
                    return ""
                return f"{h}:{m}:{s}"
            return ""

        df_sena['DURATION'] = df_source['Duration'].apply(convert_duration)

        # Membres et albums
        df_sena['NO_BAND_MEMBERS*'] = ""  # à remplir si info disponible
        df_sena['ALBUM_TITLE'] = df_source['ALBUM TITLE'] 
        df_sena['NO_ALBUM_TRACKS'] = ""  # vide
        df_sena['RECORDING_YEAR*'] = df_source['YEAR OF RECORDING']
        # Map des colonnes
        df_sena['RECORDING_COUNTRY'] = df_source['COUNTRY OF RECORDING'].map(pays_codes).fillna("")
        df_sena['CTY_CODE_PRODUCER'] = df_source['LABEL COUNTRY'].map(pays_codes).fillna("")


        df_sena['FIRST_RELEASE_YEAR*'] = df_source['YEAR OF RECORDING']
        df_sena['PROOF_URL'] = df_source['PROOF (URL link)']

        # Réordonner colonnes pour correspondre au format exact demandé
        df_sena = df_sena[['SENA_NUMBER*', 'ROLE*', 'ARTIST*', 'TITLE*', 'VERSION', 'ISRC', 'TRACKID',
                            'INSTRUMENTS', 'NO_OF_INSTRUMENTS', 'CATEGORY*', 'LABEL', 'GENRE*', 'COMPOSER',
                            'DURATION', 'NO_BAND_MEMBERS*', 'ALBUM_TITLE', 'NO_ALBUM_TRACKS', 'RECORDING_YEAR*',
                            'RECORDING_COUNTRY', 'FIRST_RELEASE_YEAR*', 'CTY_CODE_PRODUCER', 'PROOF_URL']]




        #AIE


        # --- Colonnes dans le bon ordre ---
        colonnes_aie = [
            "Stage Name",
            "Function",
            "Featuring as a Session Musician (Y/N)",
            "Total No. of Main Artists on Track",
            "% Royalty Share",
            "Track Title",
            "Title Version (Video, Remix, Live, etc)",
            "Main Artists (Separate with semi-colons '/')",
            "Duration (hh24:mi:ss)",
            'COUNTRY OF RECORDING',
            "Original Release Year",
            "Category",
            "ISRC Code",
            "Album Title",
            "Type of Media",
            "Catalogue Number",
            "Country of Origin of the Label",
            "Recording Year",
            "Record Label"
        ]
        df_aie = pd.DataFrame()


        # 1️⃣ Stage Name
        df_aie["Stage Name"] = df_source['ALIAS']

        # 2️⃣ Function = Premier instrument
        df_aie["Function"] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: str(x).split(",")[0].strip() if pd.notnull(x) else ""
        )

        # 3️⃣ Featuring as a Session Musician (Y/N)
        df_aie["Featuring as a Session Musician (Y/N)"] = df_source["ROLE"].apply(
            lambda x: "N" if str(x).upper() == "NFA" else ("Y" if str(x).upper() == "FA" else "")
        )


        # Remplacer séparateurs et nettoyer — version simple et robuste
        def clean_main_artists_simple(val):
            if pd.isna(val):
                return ""
            s = str(val)

            # Remplacements simples (ordre non sensible)
            replacements = [" feat ", "feat ", "feat." ,"featuring ", " & ", "&", " et ", "et", ",", ";"]
            for r in replacements:
                s = s.replace(r, "/")

            # Remplacer doublons de slash éventuels et retirer espaces inutiles
            while "//" in s:
                s = s.replace("//", "/")
            parts = [p.strip() for p in s.split("/") if p.strip()]
            return "/".join(parts)

        # Appliquer pour remplir la colonne "Main Artists"
        df_aie["Main Artists (Separate with semi-colons '/')"] = df_source['RELEASE ARTIST / GROUP'].apply(clean_main_artists_simple)

        # Compter le nombre d'artistes (séparés par '/')
        def count_artists_simple(s):
            if not s:
                return 0
            return len(s.split("/"))

        df_aie["Total No. of Main Artists on Track"] = df_aie["Main Artists (Separate with semi-colons '/')"].apply(count_artists_simple)


        # 5️⃣ % Royalty Share (vide par défaut)
        df_aie["% Royalty Share"] = df_source['ROYALTY SHARE']

        # 6️⃣ Track Title
        df_aie["Track Title"] = df_source['TRACK TITLE']

        # 7️⃣ Title Version (Video, Remix, Live, etc) (vide)
        df_aie["Title Version (Video, Remix, Live, etc)"] = ""

        # 9️⃣ Duration (hh24:mi:ss)
        df_aie["Duration (hh24:mi:ss)"] = df_source['Duration']

        # 🔟 Country of Recording (avec conversion code pays si dispo)
        def get_country_code(pays):
            pays = str(pays).strip()
            return pays_codes.get(pays, pays)

        df_aie['COUNTRY OF RECORDING'] = df_source['COUNTRY OF RECORDING'].apply(get_country_code)

        # 11️⃣ Original Release Year
        df_aie["Original Release Year"] = df_source['YEAR OF RECORDING']

        # 12️⃣ Category (toujours M)
        df_aie["Category"] = ""

        # 13️⃣ ISRC Code
        df_aie["ISRC Code"] = df_source['ISRC CODE']

        # 14️⃣ Album Title
        df_aie["Album Title"] = df_source['ALBUM TITLE']

        # 15️⃣ Type of Media
        df_aie["Type of Media"] = ""

        # 16️⃣ Catalogue Number (vide)
        df_aie["Catalogue Number"] = df_source["UPC"].astype("string")

        # 17️⃣ Country of Origin of the Label (code pays)
        df_aie["Country of Origin of the Label"] = df_source['LABEL COUNTRY'].apply(get_country_code)

        # 18️⃣ Recording Year
        df_aie["Recording Year"] = df_source['YEAR OF RECORDING']

        # 19️⃣ Record Label
        df_aie["Record Label"] = df_source['LABEL NAME']

        # --- Colonnes dans le bon ordre ---
        colonnes_aie = [
            "Stage Name",
            "Function",
            "Featuring as a Session Musician (Y/N)",
            "Total No. of Main Artists on Track",
            "% Royalty Share",
            "Track Title",
            "Title Version (Video, Remix, Live, etc)",
            "Main Artists (Separate with semi-colons '/')",
            "Duration (hh24:mi:ss)",
            'COUNTRY OF RECORDING',
            "Original Release Year",
            "Category",
            "ISRC Code",
            "Album Title",
            "Type of Media",
            "Catalogue Number",
            "Country of Origin of the Label",
            "Recording Year",
            "Record Label"
        ]

        df_aie = df_aie[colonnes_aie]

        # Artisti

        df_artisti = pd.DataFrame()

        # 1️⃣ Mapping des colonnes
        df_artisti["Album Title"] = df_source["ALBUM TITLE"]
        df_artisti["Album Artist"] = df_source["RELEASE ARTIST / GROUP"]
        df_artisti["Track Number"] = ""  # Pas dans df_source
        df_artisti["Track Title"] = df_source["TRACK TITLE"]
        df_artisti["Track Version"] = df_source["Version"]
        df_artisti["Track Main Artist"] = df_source["ARTIST NAME"]
        df_artisti["ISRC"] = df_source["ISRC CODE"]
        df_artisti["Agent Track ID"] = ""  # Si tu as un ID interne, tu peux le mettre ici
        df_artisti["First Maker Name(s)"] = ""
        df_artisti["First Maker Nationality(ies)"] = df_source["LABEL COUNTRY"]
        df_artisti[" Country of Fixation"] = df_source["COUNTRY OF RECORDING"]

        # Conversion des années
        df_artisti["Year of Fixation"] = df_source["YEAR OF RECORDING"].apply(
            lambda x: str(int(x)) if pd.notnull(x) and str(x).strip() else ""
        )
        df_artisti["Year of Release"] = df_artisti["Year of Fixation"]  # même valeur si pas de champ séparé

        df_artisti["Label"] = df_source["LABEL NAME"]
        df_artisti["Media Type"] = df_source["RELEASE FORMAT"]
        df_artisti[" UPC Number"] = df_source["UPC"].astype("string")
        df_artisti["Catalog Number"] = ""  # pas dans df_source
        df_artisti["Genre"] = df_source["RELEASE TYPE"]
        df_artisti["Track Duration "] = df_source["Duration"]
        df_artisti["Composer"] = ""  # pas dans df_source
        df_artisti["Track Proof"] = df_source["PROOF (URL link)"]

        # 2️⃣ Réordonner les colonnes selon la structure Artisti
        artisti_cols = [
            'Album Title',
            'Album Artist',
            'Track Number',
            'Track Title',
            'Track Version',
            'Track Main Artist',
            'ISRC',
            'Agent Track ID',
            'First Maker Name(s)',
            'First Maker Nationality(ies)',
            ' Country of Fixation',
            'Year of Fixation',
            'Year of Release',
            'Label',
            'Media Type',
            ' UPC Number',
            'Catalog Number',
            'Genre',
            'Track Duration ',
            'Composer',
            'Track Proof'
        ]
        df_artisti = df_artisti[artisti_cols]

        #ARTISTI_2

        df_artisti_2 = pd.DataFrame()

        df_artisti_2["Track Number"] = ""  # Pas dans df_source
        df_artisti_2["Track Title"] = df_source["TRACK TITLE"]
        df_artisti_2["Track Version"] = df_source["Version"]
        df_artisti_2["Track Main Artist"] = df_source["ARTIST NAME"]
        df_artisti_2["ISRC"] = df_source["ISRC CODE"]
        df_artisti_2["Agent Track ID"] = ""  # Si tu as un ID interne, tu peux le mettre ici
        df_artisti_2["Performer First Name"] = df_source["ARTIST NAME"].apply(lambda x: str(x).split()[0] if pd.notnull(x) and str(x).strip() and str(x).strip().lower() != "nan" else "")
        df_artisti_2["Performer Last Name"] = df_source["ARTIST NAME"].apply(lambda x: " ".join(str(x).split()[1:]) if pd.notnull(x) and str(x).strip() and str(x).strip().lower() != "nan" else "")
        df_artisti_2["Performer Date Of Birth"] = ""  # pas dans df_source
        df_artisti_2["Performer IPN"] = ""  # pas dans df_source
        df_artisti_2["Performer Citizenship"] = ""
        df_artisti_2[" Stage First Name"] = ""
        df_artisti_2[" Stage Last Name"] = ""
        df_artisti_2["Artisti Performer ID "] = ""  # pas dans df_source
        df_artisti_2["Agent Performer ID "] = ""  # pas dans df_source 
        df_artisti_2["Performer Category"] = df_source["ROLE"]
        df_artisti_2["Role"] = df_source["INSTRUMENT(S) / VOCALS"].apply(
            lambda x: "vocalist" if pd.notnull(x) and "vocal" in str(x).lower() else "instrumentalist"
        )
        df_artisti_2["Instruments"] = df_source["INSTRUMENT(S) / VOCALS"]
        df_artisti_2["Total No. Of Featured Artists (leave blank if unknown)"] = ""  # à remplir si info disponible
        df_artisti_2["Total No. Of Non-Featured Artists (leave blank if unknown)"] = ""  # à remplir si info disponible


        artisti_cols_2 = ['Track Number',
        'Track Title',
        'Track Version',
        'Track Main Artist',
        'ISRC',
        'Agent Track ID',
        'Performer First Name',
        'Performer Last Name',
        'Performer Date Of Birth',
        'Performer IPN',
        'Performer Citizenship',
        ' Stage First Name',
        ' Stage Last Name',
        'Artisti Performer ID ',
        'Agent Performer ID ',
        'Performer Category',
        'Role',
        'Instruments',
        'Total No. Of Featured Artists (leave blank if unknown)',
        'Total No. Of Non-Featured Artists (leave blank if unknown)']

        df_artisti_2 = df_artisti_2[artisti_cols_2]

        #EJI

        df_eji = pd.DataFrame()

        df_eji['Title of the track'] = df_source['TRACK TITLE']
        df_eji['Title of the album'] = df_source['ALBUM TITLE']
        df_eji['Name of the main artist or band under whose name this track was released'] = df_source['RELEASE ARTIST / GROUP']

        # Soliste ou membre de groupe - même logique que is_main_artist_partial
        def is_soloist(row):
            main = str(row.get('RELEASE ARTIST / GROUP', '')).strip().lower()
            pseudos = [p.strip().lower() for p in str(row.get('ALIAS', '')).split(';') if p.strip()]
            rep = str(row.get('ARTIST NAME', '')).strip().lower()
            if any(p and p in main for p in pseudos + [rep]):
                return 'soloist'
            return 'band member'

        df_eji['Mark here if you were a soloist or a band member on this track'] = df_source.apply(is_soloist, axis=1)
        df_eji['Number of band members'] = ''
        # Type de contribution selon les instruments/voix
        vocal_keywords = ['vocal', 'singer', 'soprano', 'alto', 'tenor', 'baritone',
                        'mezzo', 'falsetto', 'rap', 'yodel', 'whistling', 'contralto']
        
        df_eji['conductor / choir leader'] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: 'X' if pd.notnull(x) and 'conductor' in str(x).lower() else ''
        )
        df_eji['instrumentalist'] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: 'X' if pd.notnull(x) and not any(k in str(x).lower() for k in vocal_keywords + ['spoken word']) else ''
        )
        df_eji['singer'] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: 'X' if pd.notnull(x) and any(k in str(x).lower() for k in vocal_keywords) else ''
        )
        df_eji['speaker / reader / vocal artist '] = df_source['INSTRUMENT(S) / VOCALS'].apply(
            lambda x: 'X' if pd.notnull(x) and 'spoken word' in str(x).lower() else ''
        )
        df_eji['Record label'] = df_source['LABEL NAME']

        df_eji['Year of release'] = df_source['YEAR OF RECORDING']

        df_eji['ISRC'] = df_source['ISRC CODE']


        # ------------------------- EXPORT EXCEL -------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            if "SPED" in selected_tabs:
                df_sped.to_excel(writer, sheet_name="SPED", index=False)
            if "Playright" in selected_tabs:
                df_playright.to_excel(writer, sheet_name="Playright", index=False)
            if "SwissPerf" in selected_tabs:
                df_swissperf.to_excel(writer, sheet_name="SwissPerf", index=False)
            if "SENA" in selected_tabs:
                df_sena.to_excel(writer, sheet_name="SENA", index=False)
            if "AIE" in selected_tabs:
                df_aie.to_excel(writer,sheet_name="AIE", index=False)
            if "Artisti" in selected_tabs:
                df_artisti.to_excel(writer, sheet_name="Artisti", index=False)
            if "Artisti_2" in selected_tabs:
                df_artisti_2.to_excel(writer, sheet_name="Artisti_2", index=False)
            if "EJI" in selected_tabs:
                df_eji.to_excel(writer, sheet_name="EJI", index=False)

        st.download_button(
        label="⬇️ Télécharger le fichier Excel",
        data=output.getvalue(),
        file_name="export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

        st.success("✅ Fichier prêt au téléchargement !")

    except Exception as e:
        st.error(f"Une erreur est survenue : {str(e)}")
