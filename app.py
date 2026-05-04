import streamlit as st
import os

st.set_page_config(page_title="Tabellini Live", layout="centered")
st.title("🏆 I Tuoi Tabellini Aggiornati")

CARTELLA_FOTO = "foto_tabellini"

if os.path.exists(CARTELLA_FOTO):
    # Legge tutte le foto nella cartella
    file_immagini = [f for f in os.listdir(CARTELLA_FOTO) if f.endswith(".png")]
    nomi_giocatori = [f.replace(".png", "") for f in file_immagini]
    
    if nomi_giocatori:
        scelta = st.selectbox("Seleziona il tuo nome per vedere il tabellino:", sorted(nomi_giocatori))
        
        # Mostra la foto corrispondente
        percorso_foto = os.path.join(CARTELLA_FOTO, f"{scelta}.png")
        st.image(percorso_foto, use_container_width=True)
    else:
        st.warning("Nessun tabellino trovato. L'amministratore deve ancora generare le foto.")
else:
    st.error("Cartella foto non trovata!")