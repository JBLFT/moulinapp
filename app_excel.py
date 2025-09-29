import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Convertisseur Excel", layout="wide")
st.title("Convertisseur Excel Multi-Onglets")
st.markdown("""
Cette application permet de transformer votre fichier Excel source en quatre onglets :  
**SPED**, **Playright**, **SwissPerf**, **SENA**.
""")

# Upload du fichier Excel
uploaded_file = st.file_uploader("Choisissez un fichier Excel source", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Lecture du fichier Excel
        df_source = pd.read_excel(uploaded_file, header=1)

        # ------------------------- LOGIQUE SPED -------------------------
        df_source["Duration (HH:MM:SS)"] = df_source["Duration (HH:MM:SS)"].astype(str)
        df_source["Duration (HH:MM:SS)"] = df_source["Duration (HH:MM:SS)"].apply(
            lambda x: str(int(x.split(":")[0])) + ":" + x.split(":")[1] if ":" in x else x
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
        df_sped["Artiste représenté\nRepresented artist"] = df_source['Artiste représenté\nRepresented artist']

        df_sped['Pseudonyme\nStage Name'] = df_source['Pseudonyme\nStage Name']

        # Premier instrument = premier mot des instruments
        df_sped["Premier instrument\nFirst Function on the title"] = df_source["Instruments / Vocals"].apply(
            lambda x: str(x).split(",")[0] if pd.notnull(x) else ""
        )

        # Autres instruments = tout sauf le premier (après la 1ère virgule)
        df_sped["Autres instruments\nOther  Functions on the title"] = df_source["Instruments / Vocals"].apply(
            lambda x: ",".join([i.strip() for i in str(x).split(",")[1:]]) if pd.notnull(x) and "," in str(x) else ""
        )

        # Statut Artiste Principal et Membre du groupe 
        def is_main_artist_partial(row):
            main = str(row.get("Main artist / group", "")).strip().lower()
            pseudos = str(row.get("Pseudonyme\nStage Name", "")).strip().lower().split(";")
            pseudos = [p.strip() for p in pseudos if p.strip()]
            rep = str(row.get("Artiste représenté\nRepresented artist", "")).strip().lower()
            
            # Vérifie si au moins un pseudo ou l'artiste représenté est dans main
            for p in pseudos + [rep]:
                if p and p in main:
                    return "yes"
            return "no"

        df_sped["Statut Artiste Principal\nMain artist\n(yes or no)"] = df_source.apply(is_main_artist_partial, axis=1)


        # Membre du groupe (vide par défaut)
        df_sped["Membre du groupe\nGroup member"] = ""

        # Nom de l'artiste principal
        df_sped["Nom de l'artiste principal\nName of Main Artist"] = df_source["Main artist / group"]

        # Nom du Groupe principal (vide par défaut)
        df_sped["Nom du Groupe principal\nName of Main Group"] = ""

        # Release Title = Release type
        df_sped["Titre général\nRelease Title"] = df_source["Track title"]

        # Titre de l'enregistrement
        df_sped["Titre de l'enregistrement\nTrack Title"] = df_source["Track title"]

        # Durée en Heure, Minute, Seconde
        # Gestion de la durée (simple)
        heures, minutes, secondes, durees = [], [], [], []
        for duree in df_source["Duration (HH:MM:SS)"]:
            d = str(duree).strip()
            durees.append(d)
            if ":" in d:
                parts = d.split(":")
                if len(parts) == 3:   # HH:MM:SS
                    h, m, s = parts
                elif len(parts) == 2: # MM:SS
                    h = "0"
                    m, s = parts
                else:
                    h, m, s = "", "", ""
            else:
                h, m, s = "", "", ""
            heures.append(h)
            minutes.append(m)
            secondes.append(s)

        df_sped["Durée\nDuration"] = durees
        df_sped["Heure\nHour"] = heures
        df_sped["Minute\nMinute"] = minutes
        df_sped["Seconde\nSecond"] = secondes

        # Producteur Original
        df_sped["Producteur Original\nOriginal Record Label"] = df_source["Label name"]

        # Nationalité du producteur d'origine
        df_sped["Nationalité du producteur d'origine\nNationality of Producer"] = df_source["Label country"]

        # Pays d'enregistrement
        df_sped["Pays d'enregistrement\nCountry of Recording"] = df_source["Country of recording"]

        # Date d'enregistrement (vide par défaut car on n'a que l'année)
        df_sped["date d'enregistrement\ndate of recording\n(dd/mm/yyyy)"] = df_source["Year of recording"]

        # Année de première publication
        df_sped["Année de première publication\nYear of first publication\n(dd/mm/yyyy)"] = df_source["Year of recording"]

        #ISRC 
        df_sped["ISRC"] = df_source['ISRC code']
            
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

        # Colonnes obligatoires
        df_playright["Results"] = ""  # vide par défaut
        df_playright["PerformerName"] = df_source.get("Artiste représenté\nRepresented artist", "")
        df_playright["AffiliationNumber"] = ""
        df_playright["AlbumTitle"] = ""
        df_playright["TitleTrack"] = df_source.get("Track title", "")
        df_playright["MainArtist"] = df_source.get("Main artist / group", "")
        df_playright["RecordingYear"] = df_source.get("Year of recording", "")
        df_playright["ProductionName"] = df_source.get("Label name", "")
        df_playright["CountryOfProduction"] = df_source.get("Label country", "")
        df_playright["CountryOfMastering"] = df_source.get("Country of recording", "")

        # PerformerRole : main artist / featured artist (logique partielle)
        def performer_role_partial(row):
            main = str(row.get("Main artist / group", "")).strip().lower()
            pseudos = str(row.get("Pseudonyme\nStage Name", "")).strip().lower().split(";")
            pseudos = [p.strip() for p in pseudos if p.strip()]
            rep = str(row.get("Artiste représenté\nRepresented artist", "")).strip().lower()
            for p in pseudos + [rep]:
                if p and p in main:
                    return "Main Artist"
            return "Featured Artist"

        # Mapping ROLE -> PerformerRole
        role_map = {
            "NFA": "C",  # musicien de session
            "FA": "A"    # contrat d'artiste
        }

        df_playright["PerformerRole"] = df_source["ROLE (Featured artist / Non featured artist)"].apply(
            lambda x: role_map.get(str(x).strip(), "")
        )

        df_playright["PerformerRole2"] = ""  # vide par défaut

        # Instruments : découpe jusqu'à 5 instruments
        def find_instrument_family(instr):
            if not instr:
                return ""
            for family, instruments in instru_dict.items():
                if instr.strip() in instruments:
                    return family
            return ""  # si non trouvé

        # Premier instrument
        df_playright["Instruments"] = df_source["Instruments / Vocals"].apply(
            lambda x: str(x).split(",")[0].strip() if pd.notnull(x) and str(x).strip() else ""
        )

        # Famille correspondante
        df_playright["InstrumentType"] = df_playright["Instruments"].apply(find_instrument_family)

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
        df_swissperf["Titre (morceau/chanson/plage)"] = df_source["Track title"]

        # Interprète principal/formation = Main artist / group
        df_swissperf["Interprète principal/formation"] = df_source["Main artist / group"]

        # Instrument(s) = Instruments / Vocals
        df_swissperf["Instrument(s)"] = df_source["Instruments / Vocals"]

        # Rôle(s)
        def map_role(role):
            role = str(role).strip().upper()
            if role == "FA":
                return "FA/S"     # Featured Artist → Soliste par défaut
            elif role == "NFA":
                return "NFA/MS"   # Non Featured Artist → Musician de studio par défaut
            else:
                return ""

        df_swissperf["Rôle(s)\nFA/S, FA/CO, FA/MF\nou\nNFA/MS,  NFA/CS, NFA/MF\net/ou\nAP"] = (
            df_source["ROLE (Featured artist / Non featured artist)"].apply(map_role)
        )
        # Pays d'enregistrement
        df_swissperf["Pays d'enregistrement"] = df_source["Country of recording"]

        # Année d'enre-gistre-ment
        df_swissperf["Année d'enre-gistre-ment"] = df_source["Year of recording"]

        # Pays de publication = Label country
        df_swissperf["Pays de publication"] = df_source["Label country"]

        # Année de publica-tion = Year of recording (par défaut)
        df_swissperf["Année de publica-tion"] = df_source["Year of recording"]

        # Album = Release type
        df_swissperf["Album"] = ""

        # Label
        df_swissperf["Label"] = df_source["Label name"]

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

        df_swissperf["Durée d'enregistrement\nen min:sec"] = df_source["Duration (HH:MM:SS)"].apply(convert_duration)

        # Réordonner les colonnes
        colonnes_swissperf = [
            'Titre (morceau/chanson/plage)',
            'Interprète principal/formation',
            'Instrument(s)',
            'Rôle(s)\nFA/S, FA/CO, FA/MF\nou\nNFA/MS,  NFA/CS, NFA/MF\net/ou\nAP',
            'Pays d\'enregistrement',
            'Année d\'enre-gistre-ment',
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
            main = str(row.get("Main artist / group", "")).strip().lower()
            pseudos = str(row.get("Pseudonyme\nStage Name", "")).strip().lower().split(";")
            pseudos = [p.strip() for p in pseudos if p.strip()]
            rep = str(row.get("Artiste représenté\nRepresented artist", "")).strip().lower()

            # Vérifie si au moins un pseudo ou l'artiste représenté est dans main
            for p in pseudos + [rep]:
                if p and p in main:
                    return "MA"
            return "S"

        df_sena['ROLE*'] = df_source.apply(map_role_sena, axis=1)
        df_sena['ARTIST*'] = df_source['Main artist / group']
        df_sena['TITLE*'] = df_source['Track title']
        df_sena['VERSION'] = df_source['Track version (Radio Edit / Extended Mix / …)']
        df_sena['ISRC'] = df_source['ISRC code']
        df_sena['TRACKID'] = ""  # vide par défaut

        # Instruments
        # Instrument principal = premier instrument
        df_sena['INSTRUMENTS'] = df_source['Instruments / Vocals'].apply(lambda x: str(x).split(",")[0] if pd.notnull(x) else "")
        # Nombre d'instruments
        df_sena['NO_OF_INSTRUMENTS'] = df_source['Instruments / Vocals'].apply(lambda x: len(str(x).split(",")) if pd.notnull(x) and str(x).strip() else 0)

        # Catégorie selon ROLE
        df_sena['CATEGORY*'] = "M"

        # Label et genre
        df_sena['LABEL'] = df_source['Label name']
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

        df_sena['DURATION'] = df_source['Duration (HH:MM:SS)'].apply(convert_duration)

        # Membres et albums
        df_sena['NO_BAND_MEMBERS*'] = ""  # à remplir si info disponible
        df_sena['ALBUM_TITLE'] = df_source['Track title']  # ou Release Title si disponible
        df_sena['NO_ALBUM_TRACKS'] = ""  # vide
        df_sena['RECORDING_YEAR*'] = df_source['Year of recording']
        # Map des colonnes
        df_sena['RECORDING_COUNTRY'] = df_source['Country of recording'].map(pays_codes).fillna("")
        df_sena['CTY_CODE_PRODUCER'] = df_source['Label country'].map(pays_codes).fillna("")


        df_sena['FIRST_RELEASE_YEAR*'] = df_source['Year of recording']
        df_sena['PROOF_URL'] = df_source['PROOF (lien URL)']

        # Réordonner colonnes pour correspondre au format exact demandé
        df_sena = df_sena[['SENA_NUMBER*', 'ROLE*', 'ARTIST*', 'TITLE*', 'VERSION', 'ISRC', 'TRACKID',
                            'INSTRUMENTS', 'NO_OF_INSTRUMENTS', 'CATEGORY*', 'LABEL', 'GENRE*', 'COMPOSER',
                            'DURATION', 'NO_BAND_MEMBERS*', 'ALBUM_TITLE', 'NO_ALBUM_TRACKS', 'RECORDING_YEAR*',
                            'RECORDING_COUNTRY', 'FIRST_RELEASE_YEAR*', 'CTY_CODE_PRODUCER', 'PROOF_URL']]


    

        # ------------------------- EXPORT EXCEL -------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_sped.to_excel(writer, sheet_name="SPED", index=False)
            df_playright.to_excel(writer, sheet_name="Playright", index=False)
            df_swissperf.to_excel(writer, sheet_name="SwissPerf", index=False)
            df_sena.to_excel(writer, sheet_name="sena", index=False)
        output.seek(0)

        st.download_button(
            label="Télécharger le fichier Excel généré",
            data=output,
            file_name="exports_multi_onglets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ Fichier prêt au téléchargement !")

    except Exception as e:
        st.error(f"Une erreur est survenue : {str(e)}")
