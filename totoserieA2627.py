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
    """Legge un pezzo specifico di un foglio Excel e ne pulisce le intestazioni."""
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
        df.columns = [str(col).split('.')[0].strip() for col in df.columns]
        return df
    except Exception as e:
        return None

# --- MENU LATERALE ---
with st.sidebar:
    st.title("Menu")
    sezione_scelta = st.radio(
        "Vai a:", 
        [
            "📝 Tabelle Pronostici", 
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
                            # Se manca qualcuno, blocca la tabella e mostra il warning
                            testo_mancanti = "\n".join([f"* {nome}" for nome in mancanti])
                            st.warning(f"🔒 **Tabellone bloccato.** Le giocate saranno visibili solo quando tutti avranno inviato la colonna.\n\nAll'appello mancano ancora **{len(mancanti)}** partecipanti:\n\n{testo_mancanti}")
                        else:
                            # Se non manca nessuno, mostra successo e sblocca la tabella
                            st.success("Tutti i partecipanti hanno inviato la colonna per questa giornata!")
                            st.dataframe(df_vista.style.apply(colora_partita_jolly, axis=1), hide_index=True, use_container_width=True)
                    else:
                        # Sistema di sicurezza: se per caso l'Excel ha problemi, mostra la tabella per non rompere l'app
                        st.info("Impossibile caricare l'elenco ufficiale dall'Excel per verificare i mancanti.")
                        st.dataframe(df_vista.style.apply(colora_partita_jolly, axis=1), hide_index=True, use_container_width=True)
                        
                else:
                    st.warning("Il foglio non è nel formato previsto.")
            else:
                st.info("Il tabellone relativo a questa giornata non è ancora disponibile.")
        else:
            st.error("Impossibile caricare i dati dei pronostici da Google Sheets.")

else:
    st.error("Errore nel caricamento del file master Excel.")
