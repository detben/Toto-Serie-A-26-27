import streamlit as st
import pandas as pd
import requests
import io
import base64

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Toto Amici - Classifiche", page_icon="🏆", layout="centered")
GIORNATA_CORRENTE = 5

# ==========================================
# 1. LINK ONEDRIVE (Per le Classifiche)
# ==========================================
ONEDRIVE_LINK = "https://1drv.ms/x/c/37257a5e51e01cb8/IQBsF-NEdoduQI8-Jzl1AwRXAQvVKRsPJTBNjnoapoL8BNY?e=Blub1N"

# ==========================================
# 2. LINK GOOGLE SHEETS (Per i Pronostici)
# ==========================================
GOOGLE_LINK = "https://docs.google.com/spreadsheets/d/1Qzd-5U0ixS5dnwEXaTovX8XpVLoentoS1cfZL3Hg8p0/edit?usp=sharing"

def converti_link_onedrive(link: str) -> str:
    """Forza il parametro di download sul link fornito."""
    link = link.strip()
    if "?" in link:
        return link.split("?")[0] + "?download=1"
    return link + "?download=1"

@st.cache_data(ttl=60)
def scarica_file_excel(link):
    """Scarica l'Excel da OneDrive."""
    try:
        download_url = converti_link_onedrive(link)
        response = requests.get(download_url)
        
        if response.status_code != 200:
            b64 = base64.b64encode(link.encode("utf-8")).decode("utf-8").rstrip("=").replace("/", "_").replace("+", "-")
            api_url = f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"
            response = requests.get(api_url)
            
        if response.status_code == 200:
            return io.BytesIO(response.content)
        else:
            st.error(f"Accesso negato da OneDrive (Errore {response.status_code}).")
            return None
    except Exception as e:
        st.error(f"Errore di connessione OneDrive: {e}")
        return None

@st.cache_data(ttl=60)
def scarica_file_google(link):
    """Scarica il file da Google Sheets convertendolo al volo in Excel."""
    try:
        if "/edit" in link:
            download_url = link.split("/edit")[0] + "/export?format=xlsx"
        else:
            download_url = link
            
        response = requests.get(download_url)
        
        if response.status_code == 200:
            return io.BytesIO(response.content)
        else:
            st.error(f"Accesso negato a Google Sheets (Errore {response.status_code}). Assicurati che sia su 'Chiunque abbia il link'.")
            return None
    except Exception as e:
        st.error(f"Errore di connessione a Google Sheets: {e}")
        return None

def leggi_sezione_classifica(excel_bytes, nome_foglio, riga_inizio, num_righe, colonne):
    """
    Legge un pezzo specifico di un foglio Excel e ne pulisce le intestazioni.
    """
    try:
        excel_bytes.seek(0) # Riporta il cursore a zero
        df = pd.read_excel(
            excel_bytes, 
            sheet_name=nome_foglio,
            skiprows=riga_inizio - 1, 
            nrows=num_righe,
            usecols=colonne
        )
        df = df.dropna(how='all')
        
        # TRUCCO ANTICRASH: Taglia i numeretti (.1, .2) ma aggiunge uno spazio invisibile
        # se il nome esiste già, così Streamlit non va in errore per i doppioni.
        nuove_colonne = []
        for col in df.columns:
            nome_pulito = str(col).split('.')[0].strip()
            while nome_pulito in nuove_colonne:
                nome_pulito += " " # Aggiunge spazio invisibile
            nuove_colonne.append(nome_pulito)
            
        df.columns = nuove_colonne
        return df
    except Exception as e:
        return None

# --- MENU LATERALE ---
with st.sidebar:
    st.title("Menu")
    sezione_scelta = st.radio(
        "Vai a:", 
        [
            "👤 Area Personale",
            "📝 Tabelle Pronostici",
            "🔢 Tabelle Punteggi",
            "🗓️ Classifiche di Giornata", 
            "📊 Classifiche Trimestrali", 
            "🥇 Classifica Generale",
            "🎯 Classifiche risultati esatti e pronostici"
        ]
    )

# --- INTERFACCIA STREAMLIT ---
st.title("🏆 Toto Serie A 2026/27")
st.write("Dati aggiornati in tempo reale")
st.divider()

# Scarica entrambi i file in background
excel_file = scarica_file_excel(ONEDRIVE_LINK)
google_file = scarica_file_google(GOOGLE_LINK)

if excel_file is not None:

    # ---------------------------------------------
    # 0. AREA PERSONALE (DASHBOARD GIOCATORE)
    # ---------------------------------------------
    if sezione_scelta == "👤 Area Personale":
        st.subheader("👤 La tua Area Personale")
        st.write("Seleziona il tuo nome per vedere le tue statistiche e giocate in tempo reale.")
        
        # 1. Recupera la lista di tutti i giocatori dalla Classifica Generale (Excel)
        df_nomi = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 67, "I:I")
        
        if df_nomi is not None and not df_nomi.empty:
            lista_giocatori = sorted(df_nomi.iloc[:, 0].dropna().astype(str).str.strip().tolist())
            
            giocatore_scelto = st.selectbox("🔍 Cerca il tuo nome:", ["-- Seleziona --"] + lista_giocatori)
            
            if giocatore_scelto != "-- Seleziona --":
                st.divider()
                st.markdown(f"### Ciao, {giocatore_scelto}! 👋")
                
                # Inizializza le variabili
                pos_gen, pos_trim, pos_giorn = "-", "-", "-"
                punti_totali = "-"
                
                # --- A1. POSIZIONE GENERALE ---
                df_gen = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 67, "H:L")
                if df_gen is not None and not df_gen.empty:
                    df_gen.columns = [str(c).split('.')[0].strip() for c in df_gen.columns]
                    col_nome = df_gen.columns[1] 
                    df_gen[col_nome] = df_gen[col_nome].astype(str).str.strip()
                    dati_gen = df_gen[df_gen[col_nome] == giocatore_scelto]
                    if not dati_gen.empty:
                        pos_gen = str(dati_gen.iloc[0, 0]).replace(".0", "")
                        punti_totali = str(dati_gen.iloc[0, 2]).replace(".0", "")

                # --- A2. POSIZIONE INTERMEDIA (Trimestrale Dinamica) ---
                if GIORNATA_CORRENTE <= 14:
                    col_trim, nome_trim = "H:L", "1° Trimestre"
                elif GIORNATA_CORRENTE <= 26:
                    col_trim, nome_trim = "X:AB", "2° Trimestre"
                else:
                    col_trim, nome_trim = "AN:AR", "3° Trimestre"
                    
                df_trim = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 67, col_trim)
                if df_trim is not None and not df_trim.empty:
                    df_trim.columns = [str(c).split('.')[0].strip() for c in df_trim.columns]
                    col_nome_trim = df_trim.columns[1]
                    df_trim[col_nome_trim] = df_trim[col_nome_trim].astype(str).str.strip()
                    dati_trim = df_trim[df_trim[col_nome_trim] == giocatore_scelto]
                    if not dati_trim.empty:
                        pos_trim = str(dati_trim.iloc[0, 0]).replace(".0", "")

                # --- A3. POSIZIONE DI GIORNATA ---
                col_partenza_giorn = 2 + ((GIORNATA_CORRENTE - 1) * 15)
                colonne_giorn = [col_partenza_giorn, col_partenza_giorn + 1, col_partenza_giorn + 2, col_partenza_giorn + 3]
                df_giorn = leggi_sezione_classifica(excel_file, "Classifiche giornata", 6, 67, colonne_giorn)
                if df_giorn is not None and not df_giorn.empty:
                    df_giorn.columns = [str(c).split('.')[0].strip() for c in df_giorn.columns]
                    col_nome_giorn = df_giorn.columns[1]
                    df_giorn[col_nome_giorn] = df_giorn[col_nome_giorn].astype(str).str.strip()
                    dati_giorn = df_giorn[df_giorn[col_nome_giorn] == giocatore_scelto]
                    if not dati_giorn.empty:
                        pos_giorn = str(dati_giorn.iloc[0, 0]).replace(".0", "")

                # Stampa le Metriche in riga
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("🥇 Pos. Generale", f"{pos_gen}°")
                col2.metric(f"📊 Pos. Intermedia ({nome_trim})", f"{pos_trim}°")
                col3.metric(f"🗓️ Pos. Giornata {GIORNATA_CORRENTE}", f"{pos_giorn}°")
                col4.metric("⭐ Punti Totali", punti_totali)
                
                st.divider()

                # --- B. LE GIOCATE EFFETTUATE (Da GOOGLE SHEETS) ---
                st.markdown(f"#### 📝 Le tue Giocate (Giornata {GIORNATA_CORRENTE})")
                
                if google_file is not None:
                    nome_foglio_pronostici = f"Giornata {GIORNATA_CORRENTE}"
                    try:
                        google_file.seek(0)
                        df_pron = pd.read_excel(google_file, sheet_name=nome_foglio_pronostici)
                        foglio_esiste = True
                    except ValueError:
                        foglio_esiste = False
                        
                    if foglio_esiste and not df_pron.empty:
                        df_pron = df_pron.dropna(how='all').dropna(how='all', axis=1)
                        df_pron.columns = [str(c).split('.')[0].strip() for c in df_pron.columns]
                        
                        if len(df_pron.columns) >= 3:
                            col_partecipante = df_pron.columns[1]
                            col_jolly = df_pron.columns[-1]
                            
                            # Filtra i duplicati e prepara il nome
                            df_pron = df_pron.drop_duplicates(subset=[col_partecipante], keep='last')
                            df_pron[col_partecipante] = df_pron[col_partecipante].astype(str).str.strip()
                            
                            # Isola il giocatore scelto
                            dati_pron = df_pron[df_pron[col_partecipante] == giocatore_scelto]
                            
                            if not dati_pron.empty:
                                partita_scelta = str(dati_pron.iloc[0][col_jolly]).strip().lower()
                                
                                # Esclude Timestamp e Jolly dalla vista
                                colonne_da_mostrare = list(df_pron.columns)[1:-1]
                                df_vista = dati_pron[colonne_da_mostrare].copy()
                                
                                # Applica il colore rosso al Jolly
                                def colora_partita_jolly(row):
                                    styles = [''] * len(row)
                                    for i, col in enumerate(row.index):
                                        if str(col).strip().lower() == partita_scelta:
                                            styles[i] = 'background-color: #ff4b4b; color: white; font-weight: bold;'
                                    return styles
                                
                                st.dataframe(df_vista.style.apply(colora_partita_jolly, axis=1), hide_index=True, use_container_width=True)
                            else:
                                st.info("Non hai ancora inviato i pronostici per questa giornata.")
                        else:
                            st.warning("Il foglio pronostici non è nel formato previsto.")
                    else:
                        st.info(f"Il tabellone pronostici relativo alla Giornata {GIORNATA_CORRENTE} non è ancora disponibile.")
                else:
                    st.error("Impossibile caricare i dati dei pronostici da Google Sheets.")

                # --- C. DETTAGLIO PUNTEGGI (Da Excel) ---
                st.markdown(f"#### 🔢 Dettaglio Punteggi (Giornata {GIORNATA_CORRENTE})")
                
                col_partenza_punteggi = 1 + ((GIORNATA_CORRENTE - 1) * 15)
                colonne_punteggi = [col_partenza_punteggi + i for i in range(14)]
                
                df_punteggi = leggi_sezione_classifica(
                    excel_file, 
                    "Punteggi", 
                    riga_inizio=4,         
                    num_righe=67, 
                    colonne=colonne_punteggi
                )
                
                if df_punteggi is not None and not df_punteggi.empty and len(df_punteggi) > 1:
                    nuove_colonne = []
                    for i, col in enumerate(df_punteggi.columns):
                        val_intestazione = str(col).split('.')[0].strip()
                        val_riga_0 = str(df_punteggi.iloc[0, i]).strip()
                        
                        if i == 0:
                            nuovo_nome = "Partecipanti"
                        elif i >= 11: 
                            nuovo_nome = val_intestazione if val_intestazione.upper() not in ["NAN", ""] else f"Col_Speciale_{i}"
                        else:
                            nuovo_nome = f"{val_riga_0} [{val_intestazione}]"
                        
                        while nuovo_nome in nuove_colonne:
                            nuovo_nome += " "
                        nuove_colonne.append(nuovo_nome)
                        
                    df_punteggi.columns = nuove_colonne
                    df_punteggi = df_punteggi.iloc[1:].reset_index(drop=True)
                    
                    df_punteggi["Partecipanti"] = df_punteggi["Partecipanti"].astype(str).str.strip()
                    dati_giornata = df_punteggi[df_punteggi["Partecipanti"] == giocatore_scelto]
                    
                    if not dati_giornata.empty:
                        dati_giornata = (
                            dati_giornata.astype(str)
                            .replace({r'\.0$': ''}, regex=True)
                            .replace(["nan", "NaN", "None", ""], "0")
                        )
                        
                        # Colora i punteggi (1 verde chiaro, 3 verde scuro)
                        def colora_punti_partite(row):
                            styles = [''] * len(row)
                            for i, col in enumerate(row.index):
                                if 1 <= i <= 10:
                                    valore = str(row[col]).strip()
                                    if valore == "1":
                                        styles[i] = 'background-color: #b2df8a; color: black; font-weight: bold;'
                                    elif valore == "3":
                                        styles[i] = 'background-color: #33a02c; color: white; font-weight: bold;'
                            return styles

                        df_stile_giornata = dati_giornata.style.apply(colora_punti_partite, axis=1)
                        st.dataframe(df_stile_giornata, hide_index=True, use_container_width=True)
                    else:
                        st.info("I tuoi punteggi per la giornata corrente non sono ancora disponibili.")
                else:
                    st.info("La tabella punteggi della giornata corrente è in fase di aggiornamento.")
        else:
            st.error("Impossibile caricare l'elenco dei giocatori.")
    
    # ---------------------------------------------
    # 1. CLASSIFICA GENERALE (Da OneDrive)
    # ---------------------------------------------
    if sezione_scelta == "🥇 Classifica Generale":
        st.subheader("🥇 Classifica Generale")
        
        df_generale = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 65, "H:L")
        
        if df_generale is not None and not df_generale.empty:
            df_generale = df_generale.rename(columns=lambda x: "RIS. ESATTI" if "esatti" in str(x).lower() else x)
            st.dataframe(df_generale, hide_index=True, use_container_width=True, height=2325)
        else:
            st.warning("Classifica generale non trovata o formato errato.")

    # ---------------------------------------------
    # 2. CLASSIFICHE TRIMESTRALI (Da OneDrive)
    # ---------------------------------------------
    elif sezione_scelta == "📊 Classifiche Trimestrali":
        st.subheader("📊 Classifiche Trimestrali")
        tab1, tab2, tab3 = st.tabs(["1° Trimestre", "2° Trimestre", "3° Trimestre"])
        
        with tab1:
            df_trim1 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "H:L")
            if df_trim1 is not None and not df_trim1.empty:
                df_trim1 = df_trim1.rename(columns=lambda x: "RIS. ESATTI" if "esatti" in str(x).lower() else x)
                st.dataframe(df_trim1, hide_index=True, use_container_width=True, height=2330)
            else:
                st.info("Dati non disponibili")
                
        with tab2:
            df_trim2 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "X:AB")
            if df_trim2 is not None and not df_trim2.empty:
                df_trim2 = df_trim2.rename(columns=lambda x: "RIS. ESATTI" if "esatti" in str(x).lower() else x)
                st.dataframe(df_trim2, hide_index=True, use_container_width=True, height=2330)
            else:
                st.info("Dati non disponibili")

        with tab3:
            df_trim3 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "AN:AR")
            if df_trim3 is not None and not df_trim3.empty:
                df_trim3 = df_trim3.rename(columns=lambda x: "RIS. ESATTI" if "esatti" in str(x).lower() else x)
                st.dataframe(df_trim3, hide_index=True, use_container_width=True, height=2330)
            else:
                st.info("Dati non disponibili")

    # ---------------------------------------------
    # 3. CLASSIFICHE DI GIORNATA (Da OneDrive)
    # ---------------------------------------------
    elif sezione_scelta == "🗓️ Classifiche di Giornata":
        st.subheader("🗓️ Classifiche di Giornata")
        st.write("Scegli la giornata da visualizzare:")
        
        lista_giornate = [f"Giornata {i}" for i in range(3, 39)]
        indice_default = GIORNATA_CORRENTE - 3 if GIORNATA_CORRENTE >= 3 else 0
        
        giornata_scelta = st.selectbox("Seleziona", lista_giornate, index=indice_default, key="memoria_giornata_classifiche")
        
        numero_giornata = int(giornata_scelta.split()[1])
        colonna_partenza = 2 + ((numero_giornata - 1) * 15)
        colonne_giornata = [colonna_partenza, colonna_partenza + 1, colonna_partenza + 2, colonna_partenza + 3]
        
        df_giornata = leggi_sezione_classifica(excel_file, "Classifiche giornata", 6, 65, colonne_giornata)
        
        if df_giornata is not None and not df_giornata.empty:
            st.dataframe(df_giornata, hide_index=True, use_container_width=True, height=2330)
        else:
            classifica_base = []
            df_partecipanti = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 66, "I:I") 
            if df_partecipanti is not None and not df_partecipanti.empty:
                for nome in df_partecipanti.iloc[:, 0]:
                    classifica_base.append({"Posizione": 1, "Partecipanti": nome, "Punti": 0, "Risultati Esatti": 0})
                
                df_giornata_vuota = pd.DataFrame(classifica_base).sort_values(["Punti", "Partecipanti"], ascending=[False, True]).reset_index(drop=True)
                df_giornata_vuota.index += 1
                st.dataframe(df_giornata_vuota, hide_index=True, use_container_width=True)
            else:
                st.info("Classifica non ancora disponibile per questa giornata.")

    # ---------------------------------------------
    # 4. RISULTATI ESATTI E PRONOSTICI (Da OneDrive)
    # ---------------------------------------------
    elif sezione_scelta == "🎯 Classifiche risultati esatti e pronostici":
        st.subheader("🎯 Classifiche Risultati Esatti e Pronostici (1X2)")
        
        tab_esatti, tab_segni = st.tabs(["🎯 Risultati Esatti", "✅ Pronostici (1X2)"])
        
        with tab_esatti:
            df_esatti = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 65, "S:U")
            if df_esatti is not None and not df_esatti.empty:
                st.dataframe(df_esatti, hide_index=True, use_container_width=True, height=2330)
            else:
                st.info("Dati sui risultati esatti non disponibili.")
                
        with tab_segni:
            df_segni = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 65, "Y:AA")
            if df_segni is not None and not df_segni.empty:
                st.dataframe(df_segni, hide_index=True, use_container_width=True, height=2330)
            else:
                st.info("Dati sui pronostici non disponibili.")

    # ---------------------------------------------
    # 5. TABELLE PRONOSTICI (DA GOOGLE SHEETS)
    # ---------------------------------------------
    elif sezione_scelta == "📝 Tabelle Pronostici":
        st.subheader("📝 Tabelle Pronostici Inseriti")
        st.write("Scegli la giornata da controllare:")
        
        lista_giornate_pronostici = [f"Giornata {i}" for i in range(3, 39)]
        indice_default_pronostici = GIORNATA_CORRENTE - 3 if GIORNATA_CORRENTE >= 3 else 0
        
        giornata_scelta_pronostico = st.selectbox("Seleziona Giornata", lista_giornate_pronostici, index=indice_default_pronostici, key="memoria_giornata_pronostici")
        
        # Cerca i dati nel file Google
        if google_file is not None:
            nome_foglio_pronostici = giornata_scelta_pronostico
            try:
                google_file.seek(0)
                df_pron = pd.read_excel(google_file, sheet_name=nome_foglio_pronostici)
                foglio_esiste = True
            except ValueError:
                foglio_esiste = False
                
            if foglio_esiste and not df_pron.empty:
                df_pron = df_pron.dropna(how='all').dropna(how='all', axis=1)
                
                # Pulisce i nomi delle colonne
                df_pron.columns = [str(c).split('.')[0].strip() for c in df_pron.columns]
                
                # Accorcia l'intestazione lunga
                df_pron = df_pron.rename(columns=lambda x: "Elenco partecipanti" if "selezionare il proprio nome" in str(x).lower() else x)
                
                if len(df_pron.columns) >= 3:
                    col_partecipante = df_pron.columns[1]
                    col_jolly = df_pron.columns[-1]
                    
                    # ==========================================
                    # NUOVO: SISTEMA ANTIFRODE / INVII MULTIPLI
                    # ==========================================
                    conteggi = df_pron[col_partecipante].value_counts()
                    invii_multipli = conteggi[conteggi > 1]
                    
                    if not invii_multipli.empty:
                        testo_furbetti = "\n".join([f"* **{nome}** ha inviato la schedina **{volte} volte**" for nome, volte in invii_multipli.items()])
                        st.warning(f"🚨 **ATTENZIONE - INVII MULTIPLI RILEVATI:**\n\n{testo_furbetti}\n\n*Il sistema ha tenuto in considerazione solo l'ultimo invio cronologico per ciascun partecipante.*")
                    
                    # Procede con la pulizia standard (tiene solo l'ultimo invio)
                    df_pron = df_pron.drop_duplicates(subset=[col_partecipante], keep='last')
                    df_pron = df_pron.sort_values(by=col_partecipante).reset_index(drop=True)
                    
                    mappa_jolly = dict(zip(df_pron[col_partecipante], df_pron[col_jolly]))
                    
                    colonne_da_mostrare = list(df_pron.columns)[1:-1]
                    df_vista = df_pron[colonne_da_mostrare].copy()
                    
                    def colora_partita_jolly(row):
                        styles = [''] * len(row)
                        partecipante_attuale = row[col_partecipante]
                        partita_scelta = str(mappa_jolly.get(partecipante_attuale, "")).strip().lower()
                        
                        for i, col in enumerate(row.index):
                            if str(col).strip().lower() == partita_scelta:
                                styles[i] = 'background-color: #ff4b4b; color: white; font-weight: bold;'
                        return styles
                    
                    # ==========================================
                    # CONTROLLO MANCANTI E BLOCCO VISUALE TABELLA
                    # ==========================================
                    df_partecipanti = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 65, "I:I") 
                    
                    if df_partecipanti is not None and not df_partecipanti.empty:
                        tutti_i_nomi = set(df_partecipanti.iloc[:, 0].dropna().astype(str).str.strip())
                        nomi_inviati = set(df_pron[col_partecipante].dropna().astype(str).str.strip())
                        
                        mancanti = sorted(list(tutti_i_nomi - nomi_inviati))
                        
                        if mancanti:
                            testo_mancanti = "\n".join([f"* {nome}" for nome in mancanti])
                            st.warning(f"🔒 **Tabellone bloccato.** Le giocate saranno visibili solo quando tutti avranno inviato la colonna.\n\nAll'appello mancano ancora **{len(mancanti)}** partecipanti:\n\n{testo_mancanti}")
                        else:
                            st.success("Tutti i partecipanti hanno inviato la colonna per questa giornata! ")
                            st.dataframe(df_vista.style.apply(colora_partita_jolly, axis=1), hide_index=True, use_container_width=True)
                    else:
                        st.info("Impossibile caricare l'elenco ufficiale dall'Excel per verificare i mancanti.")
                        st.dataframe(df_vista.style.apply(colora_partita_jolly, axis=1), hide_index=True, use_container_width=True)
                        
                else:
                    st.warning("Il foglio non è nel formato previsto.")
            else:
                st.info("Il tabellone relativo a questa giornata non è ancora disponibile.")
        else:
            st.error("Impossibile caricare i dati dei pronostici da Google Sheets.")

    # ---------------------------------------------
    # 6. TABELLE PUNTEGGI
    # ---------------------------------------------
    elif sezione_scelta == "🔢 Tabelle Punteggi":
        st.subheader("🔢 Tabelle Punteggi")
        st.write("Scegli la giornata da visualizzare:")
        
        # Genera la lista delle giornate dalla 3 alla 38
        lista_giornate_punteggi = [f"Giornata {i}" for i in range(3, 39)]
        
        # Mantiene in memoria la giornata corrente partendo dalla 3
        indice_default_punteggi = GIORNATA_CORRENTE - 3 if GIORNATA_CORRENTE >= 3 else 0
        
        giornata_scelta_punteggi = st.selectbox(
            "Seleziona Giornata", 
            lista_giornate_punteggi, 
            index=indice_default_punteggi, 
            key="memoria_giornata_punteggi"
        )
        
        numero_giornata_punteggi = int(giornata_scelta_punteggi.split()[1])
        
        # =========================================================
        # BLOCCO LOGICO: CONTROLLO GIORNATE FUTURE
        # =========================================================
        if numero_giornata_punteggi > GIORNATA_CORRENTE:
            st.warning(f"⏳ I punteggi per la Giornata {numero_giornata_punteggi} non sono ancora disponibili perché non è ancora stata giocata.")
        else:
            # Se la giornata è uguale o precedente a quella corrente, calcola e mostra la tabella
            col_partenza_punteggi = 1 + ((numero_giornata_punteggi - 1) * 15)
            colonne_punteggi = [col_partenza_punteggi + i for i in range(14)]
            
            df_punteggi = leggi_sezione_classifica(
                excel_file, 
                "Punteggi", 
                riga_inizio=4,         
                num_righe=67, 
                colonne=colonne_punteggi
            )
            
            if df_punteggi is not None and not df_punteggi.empty and len(df_punteggi) > 1:
                try:
                    nuove_colonne = []
                    for i, col in enumerate(df_punteggi.columns):
                        val_intestazione = str(col).split('.')[0].strip()
                        val_riga_0 = str(df_punteggi.iloc[0, i]).strip()
                        
                        if i == 0:
                            nuovo_nome = "Partecipanti"
                        elif i >= 11: 
                            nuovo_nome = val_intestazione if val_intestazione.upper() not in ["NAN", ""] else f"Col_Speciale_{i}"
                        else:
                            nuovo_nome = f"{val_riga_0} [{val_intestazione}]"
                        
                        while nuovo_nome in nuove_colonne:
                            nuovo_nome += " "
                        nuove_colonne.append(nuovo_nome)
                        
                    df_punteggi.columns = nuove_colonne
                    df_punteggi = df_punteggi.iloc[1:].reset_index(drop=True)
                    
                    # 1. TRUCCO DECIMALI
                    df_punteggi = (
                        df_punteggi.astype(str)
                        .replace({r'\.0$': ''}, regex=True)
                        .replace(["nan", "NaN", "None", ""], "0")
                    )

                    # 2. LOGICA COLORI
                    def colora_punti_partite(row):
                        styles = [''] * len(row)
                        for i, col in enumerate(row.index):
                            if 1 <= i <= 10:
                                valore = str(row[col]).strip()
                                if valore == "1":
                                    styles[i] = 'background-color: #b2df8a; color: black; font-weight: bold;'
                                elif valore == "3":
                                    styles[i] = 'background-color: #33a02c; color: white; font-weight: bold;'
                        return styles

                    df_stile = df_punteggi.style.apply(colora_punti_partite, axis=1)

                    st.dataframe(df_stile, hide_index=True, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Errore tecnico durante la visualizzazione: {e}")
            else:
                st.info("La tabella punteggi relativa a questa giornata non è ancora disponibile o è in attesa di dati.")



