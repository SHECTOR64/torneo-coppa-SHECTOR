import xlwings as xw
import os
import webbrowser
import time
import tkinter as tk
from tkinter import messagebox

# --- CONFIGURAZIONE ---
NOME_FOGLIO_PRINCIPALE = "PRONO&RISULTATI"
NOME_FOGLIO_PRONO = "PRONO ORIZZ"
NOME_FOGLIO_COPPA = "CLASSIFICHE COPPA"
CARTELLA_FOTO_NOME = "foto_tabellini"
NOME_FILE_WEB = "visualizza_tabellini.html"

# Coordinate Range
RANGE_RISULTATI = "B4:T19"
RANGE_EXTRA = "B24:N38"
RANGE_CLASSIFICA = "AC3:BJ62" 
RANGE_PRONOSTICI_TUTTI = "A1:AU58"
CELLE_NERO = "AF6:AF7"
RANGE_COPPA_3 = "L175:AA220"

NOME_FOTO_RISULTATI = "risultati_partite.png"
NOME_FOTO_EXTRA = "foto_extra.png"
NOME_FOTO_CLASSIFICA = "classifica_turno.png"
NOME_FOTO_PRONO_TUTTI = "I PRONOSTICI DI TUTTI.png"
NOME_FOTO_COPPA_3 = "foto_coppa_3.png"

def scatta_foto_precise():
    cartella_script = os.path.dirname(os.path.abspath(__file__))
    os.chdir(cartella_script) 
    
    percorso_foto = os.path.join(cartella_script, CARTELLA_FOTO_NOME)
    percorso_web = os.path.join(cartella_script, NOME_FILE_WEB)

    if not os.path.exists(percorso_foto):
        os.makedirs(percorso_foto)

    try:
        try:
            wb = xw.Book.caller()
        except:
            wb = xw.Book("TURNOINCORSO.xlsm")
            
        sht_main = wb.sheets[NOME_FOGLIO_PRINCIPALE]
        sht_prono = wb.sheets[NOME_FOGLIO_PRONO]
        sht_coppa_foglio = wb.sheets[NOME_FOGLIO_COPPA]
        
        timestamp_url = int(time.time())
        timestamp_label = time.strftime('%H:%M:%S')

        # --- MODIFICA CELLE (AF6:AF7) ---
        range_nero = sht_main.range(CELLE_NERO)
        range_nero.color = (0, 0, 0)
        range_nero.api.Font.Color = 0

        # --- ESTRAZIONE FOTO ---
        sht_main.range(RANGE_RISULTATI).to_png(os.path.join(percorso_foto, NOME_FOTO_RISULTATI))
        sht_main.range(RANGE_EXTRA).to_png(os.path.join(percorso_foto, NOME_FOTO_EXTRA))
        sht_main.range(RANGE_CLASSIFICA).to_png(os.path.join(percorso_foto, NOME_FOTO_CLASSIFICA))
        sht_prono.range(RANGE_PRONOSTICI_TUTTI).to_png(os.path.join(percorso_foto, NOME_FOTO_PRONO_TUTTI))
        sht_coppa_foglio.range(RANGE_COPPA_3).to_png(os.path.join(cartella_script, NOME_FOTO_COPPA_3))
        
        # --- ESTRAZIONE TABELLINI ---
        lista_nomi = sht_main.range("AD10:AD62").value
        riga_5 = sht_main.range("5:5").value
        nomi_processati = ["I PRONOSTICI DI TUTTI"]

        for nome in lista_nomi:
            if nome:
                try:
                    nome_f = str(nome).strip()
                    col_base = riga_5.index(nome) + 1
                    range_tab = sht_main.range((5, col_base - 4), (26, col_base + 6))
                    range_tab.to_png(os.path.join(percorso_foto, f"{nome_f}.png"))
                    nomi_processati.append(nome_f)
                except:
                    continue

        # --- GENERAZIONE HTML ---
        menu_options = "".join([f'<option value="{CARTELLA_FOTO_NOME}/{n}.png?v={timestamp_url}">{n}</option>' for n in nomi_processati])
        
        STILE_TITOLO_VERDE = 'color: #00ff88; font-size: 52px; text-shadow: 3px 3px #004422; margin: 0 20px; font-family: "AR CENA", sans-serif;'
        STILE_SETTORE_VERDE = 'border: 3px solid #00ff88; background: #0a0a0a; padding: 30px; border-radius: 15px; margin-bottom: 40px;'
        BORDER_FOTO = 'border: 2px solid #00ff88;'

        sezione_coppa = f'''
            <hr style="border: 5px solid gold; margin: 60px 0;">
            <div class="section" style="border: 3px solid gold; background: #0a0a0a; padding: 30px; border-radius: 15px;">
                <h1 style="color: gold; font-size: 52px; text-shadow: 2px 2px black; margin-bottom: 40px; font-family: 'AR CENA', sans-serif;">🏆 1° Coppa SHECTOR 🏆</h1>
                
                <!-- PRIMA FASE -->
                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold; margin-bottom: 40px;">
                    <div style="display: flex; flex-direction: row; justify-content: center; align-items: flex-start; gap: 30px;">
                        <img src="immagini_fisse/Primafoto.png" style="width: 25%; max-height: 560px; object-fit: contain; border: 2px solid gold;">
                        <img src="foto_coppa_1a.png?v={timestamp_url}" style="width: 25%; height: 950px; object-fit: contain; border: 2px solid #444;">
                        <img src="foto_coppa_1b.png?v={timestamp_url}" style="width: 25%; height: 950px; object-fit: contain; border: 2px solid #444;">
                    </div>
                </div>

                <!-- SECONDA FASE -->
                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold; margin-bottom: 40px;">
                    <div style="display: flex; flex-direction: row; justify-content: center; align-items: flex-start; gap: 20px;">
                        <div style="width: 28%; display: flex; flex-direction: column; gap: 15px; align-items: center; margin: 0;">
                            <img src="immagini_fisse/2afase.png" style="width: 100%; height: auto; display: block;">
                            <img src="immagini_fisse/CLASS2AFASE.png" style="width: 75%; height: auto; border: 1px solid gold;">
                        </div>
                        <div style="width: 52%; margin: 0;">
                            <img src="foto_coppa_2.png?v={timestamp_url}" style="width: 100%; border: 2px solid #444; display: block;">
                        </div>
                    </div>
                </div>

                <!-- TERZA FASE AGGIORNATA -->
                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold;">
                    <div style="display: flex; flex-direction: row; justify-content: center; align-items: flex-start; gap: 20px;">
                        <!-- Foto sinistra ridotta -->
                        <img src="immagini_fisse/3afase.png" style="width: 30%; height: auto; display: block;">
                        <!-- Foto destra ingrandita -->
                        <img src="{NOME_FOTO_COPPA_3}?v={timestamp_url}" style="width: 65%; border: 2px solid #444; display: block;">
                    </div>
                </div>
            </div>
        '''

        html_content = f'''
        <!DOCTYPE html>
        <html lang="it">
        <head>
            <meta charset="UTF-8">
            <title>16° Torneo Pronostici</title>
            <style>
                body {{ 
                    background-color: #121212; 
                    color: white; 
                    font-family: "AR CENA", "Segoe UI", sans-serif; 
                    text-align: center; 
                    margin: 0; 
                    padding: 15px; 
                }}
                .flex-top {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 30px; max-width: 1400px; margin: 0 auto; }}
                .col-risultati {{ flex: 2; text-align: left; }}
                .col-extra {{ flex: 1; text-align: right; }}
                .img-risultati {{ width: 100%; max-width: 850px; border-radius: 8px; {BORDER_FOTO} }}
                .img-extra {{ width: 100%; max-width: 450px; border-radius: 8px; border: 1px solid #444; }}
                .img-max {{ width: 100%; max-width: none; border-radius: 5px; {BORDER_FOTO} }}
                .titolo-sezione {{ 
                    color: #00ff88; 
                    font-size: 32px; 
                    margin-bottom: 20px; 
                    display: block; 
                    font-family: "AR CENA", sans-serif;
                    text-decoration: underline;
                    text-underline-offset: 8px;
                }}
                select {{ 
                    padding: 15px; 
                    border-radius: 8px; 
                    background: #1a1a1a; 
                    color: #00ff88; 
                    border: 2px solid #00ff88; 
                    font-size: 22px; 
                    width: 90%; 
                    max-width: 600px; 
                    cursor: pointer; 
                }}
                .foto-standard {{ max-width: 850px; width: 95%; border-radius: 12px; border: 2px solid #00ff88; margin-top: 25px; }}
                .foto-wide {{ width: 100%; border-radius: 5px; border: 2px solid #00ff88; margin-top: 25px; }}
                #foto-tabellino {{ display: none; }}
                .footer-info {{ color:#00ff88; font-size:1rem; margin-top:15px; opacity: 0.7; }}
                .emoji-titolo {{ font-size: 60px; vertical-align: middle; }}
            </style>
        </head>
        <body>
            <div style="{STILE_SETTORE_VERDE}">
                <div style="display: flex; align-items: center; justify-content: center;">
                    <span class="emoji-titolo">🏆</span>
                    <h1 style="{STILE_TITOLO_VERDE}">16° Torneo PRONOSTICI SERIE A - 2025/26</h1>
                    <span class="emoji-titolo">🏆</span>
                </div>
            </div>

            <div style="{STILE_SETTORE_VERDE}">
                <span class="titolo-sezione">⚽ I RISULTATI ⚽</span>
                <div class="flex-top">
                    <div class="col-risultati"><img src="{CARTELLA_FOTO_NOME}/{NOME_FOTO_RISULTATI}?v={timestamp_url}" class="img-risultati"></div>
                    <div class="col-extra"><img src="{CARTELLA_FOTO_NOME}/{NOME_FOTO_EXTRA}?v={timestamp_url}" class="img-extra"></div>
                </div>
            </div>

            <div style="{STILE_SETTORE_VERDE}">
                <span class="titolo-sezione">📊 LE CLASSIFICHE 📊</span>
                <img src="{CARTELLA_FOTO_NOME}/{NOME_FOTO_CLASSIFICA}?v={timestamp_url}" class="img-max">
            </div>

            <div style="{STILE_SETTORE_VERDE}">
                <span class="titolo-sezione">📝 I TABELLINI 📝</span>
                <select id="userSelect" onchange="mostraFoto(this)">
                    <option value="">-- SELEZIONA GIOCATORE O PRONOSTICI --</option>
                    {menu_options}
                </select>
                <div id="container-foto">
                    <img id="foto-tabellino" src="">
                </div>
                <div class="footer-info">Ultimo aggiornamento: {timestamp_label}</div>
            </div>
            
            {sezione_coppa}
            
            <script>
                function mostraFoto(selectElement) {{
                    var url = selectElement.value;
                    var img = document.getElementById('foto-tabellino');
                    if(url === "") {{ 
                        img.style.display = "none"; 
                    }} else {{ 
                        img.src = url; 
                        img.style.display = "inline-block";
                        if(selectElement.options[selectElement.selectedIndex].text === "I PRONOSTICI DI TUTTI") {{
                            img.className = "foto-wide";
                        }} else {{
                            img.className = "foto-standard";
                        }}
                    }}
                }}
            </script>
        </body>
        </html>
        '''
        with open(percorso_web, "w", encoding="utf-8") as f:
            f.write(html_content)
        webbrowser.open(f"file://{percorso_web}")

    except Exception as e:
        root = tk.Tk(); root.withdraw(); messagebox.showerror("ERRORE", str(e)); root.destroy()

if __name__ == "__main__":
    scatta_foto_precise()