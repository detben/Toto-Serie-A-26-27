import streamlit as st
import pandas as pd
import requests
import io
import base64

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Toto Amici - Classifiche", page_icon="🏆", layout="centered")
GIORNATA_CORRENTE = 4

# ==========================================
# INCOLLA QUI IL LINK PRESO DA "CONDIVIDI"
# ==========================================
ONEDRIVE_LINK = "https://1drv.ms/x/c/37257a5e51e01cb8/IQBsF-NEdoduQI8-Jzl1AwRXAQvVKRsPJTBNjnoapoL8BNY?e=Blub1N"

def converti_link_onedrive(link: str) -> str:
    """Forza il parametro di download sul link fornito."""
    link = link.strip()
    if "?" in link:
        return link.split("?")[0] + "?download=1"
    return link + "?download=1"

@st.cache_data(ttl=60)
def scarica_file_excel(link):
    """Scarica l'Excel tentando prima il download diretto, poi l'API di OneDrive."""
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
            st.error(f"Accesso negato da OneDrive (Errore {response.status_code}). Assicurati che le impostazioni del link siano su 'Chiunque abbia il collegamento'.")
            return None
    except Exception as e:
        st.error(f"Errore di connessione: {e}")
        return None

def leggi_sezione_classifica(excel_bytes, nome_foglio, riga_inizio, num_righe, colonne):
    """
    Legge un pezzo specifico di un foglio Excel e ne pulisce le intestazioni.
    """
    try:
        df = pd.read_excel(
            excel_bytes, 
            sheet_name=nome_foglio,
            skiprows=riga_inizio - 1, 
            nrows=num_righe,
            usecols=colonne
        )
        df = df.dropna(how='all')
        
        # TRUCCO: Taglia via i numeretti (.1, .2) che Pandas aggiunge ai nomi duplicati
        df.columns = [str(col).split('.')[0] for col in df.columns]
        
        return df
    except Exception as e:
        return None

# --- MENU LATERALE ---
with st.sidebar:
    st.title("Menu")
    sezione_scelta = st.radio(
        "Vai a:", 
        [
            "🗓️ Classifiche di Giornata", 
            "📊 Classifiche Trimestrali", 
            "🥇 Classifica Generale",
            "🎯 Classifiche risultati esatti e pronostici"
        ]
    )

# --- INTERFACCIA STREAMLIT ---
st.title("🏆 Classifiche Toto Amici")
st.write("Dati aggiornati in tempo reale")
st.divider()

excel_file = scarica_file_excel(ONEDRIVE_LINK)

if excel_file is not None:
    
    # ---------------------------------------------
    # 1. CLASSIFICA GENERALE
    # ---------------------------------------------
    if sezione_scelta == "🥇 Classifica Generale":
        st.subheader("🥇 Classifica Generale")
        df_generale = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 65, "H:L")
        if df_generale is not None and not df_generale.empty:
            st.dataframe(df_generale, hide_index=True, use_container_width=True, height=700)
        else:
            st.warning("Classifica generale non trovata o formato errato.")

    # ---------------------------------------------
    # 2. CLASSIFICHE TRIMESTRALI
    # ---------------------------------------------
    elif sezione_scelta == "📊 Classifiche Trimestrali":
        st.subheader("📊 Classifiche Trimestrali")
        tab1, tab2, tab3 = st.tabs(["1° Trimestre", "2° Trimestre", "3° Trimestre"])
        
        with tab1:
            df_trim1 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "H:L")
            if df_trim1 is not None and not df_trim1.empty:
                st.dataframe(df_trim1, hide_index=True, use_container_width=True, height=700)
            else:
                st.info("Dati non disponibili")
                
        with tab2:
            df_trim2 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "X:AB")
            if df_trim2 is not None and not df_trim2.empty:
                st.dataframe(df_trim2, hide_index=True, use_container_width=True, height=700)
            else:
                st.info("Dati non disponibili")

        with tab3:
            df_trim3 = leggi_sezione_classifica(excel_file, "Classifiche trimestrali", 6, 65, "AN:AR")
            if df_trim3 is not None and not df_trim3.empty:
                st.dataframe(df_trim3, hide_index=True, use_container_width=True, height=700)
            else:
                st.info("Dati non disponibili")

    # ---------------------------------------------
    # 3. CLASSIFICHE DI GIORNATA (Orizzontali)
    # ---------------------------------------------
    elif sezione_scelta == "🗓️ Classifiche di Giornata":
        st.subheader("🗓️ Classifiche di Giornata")
        st.write("Scegli la giornata da visualizzare:")
        
        lista_giornate = [f"Giornata {i}" for i in range(3, 39)]
        
        # Calcola la posizione nella lista (Giornata 1 corrisponde all'indice 0)
        indice_default = GIORNATA_CORRENTE - 3
        
        giornata_scelta = st.selectbox(
            "Seleziona", 
            lista_giornate,
            index=indice_default,
            key="memoria_giornata_classifiche" # Salva la scelta dell'utente durante la sessione
        )
        
        numero_giornata = int(giornata_scelta.split()[1])
        colonna_partenza = 2 + ((numero_giornata - 1) * 15)
        colonne_giornata = [colonna_partenza, colonna_partenza + 1, colonna_partenza + 2, colonna_partenza + 3]
        
        df_giornata = leggi_sezione_classifica(
            excel_file, 
            "Classifiche giornata", 
            riga_inizio=6,         
            num_righe=65,          
            colonne=colonne_giornata
        )
        
        if df_giornata is not None and not df_giornata.empty:
            st.dataframe(df_giornata, hide_index=True, use_container_width=True, height=700)
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
    # 4. RISULTATI ESATTI E PRONOSTICI
    # ---------------------------------------------
    elif sezione_scelta == "🎯 Classifiche risultati esatti e pronostici":
        st.subheader("🎯 Classifiche Risultati Esatti e Pronostici (1X2)")
        st.write("Statistiche aggiornate dei partecipanti.")
        
        tab_esatti, tab_segni = st.tabs(["🎯 Risultati Esatti", "✅ Pronostici (1X2)"])
        
        with tab_esatti:
            df_esatti = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 66, "S:U")
            if df_esatti is not None and not df_esatti.empty:
                st.dataframe(df_esatti, hide_index=True, use_container_width=True, height=700)
            else:
                st.info("Dati sui risultati esatti non disponibili.")
                
        with tab_segni:
            df_segni = leggi_sezione_classifica(excel_file, "Classifica generale", 6, 66, "Y:AA")
            if df_segni is not None and not df_segni.empty:
                st.dataframe(df_segni, hide_index=True, use_container_width=True, height=700)
            else:
                st.info("Dati sui pronostici non disponibili.")

else:
    st.error("Errore nel caricamento del file master.")
