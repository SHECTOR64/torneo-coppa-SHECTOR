import xlwings as xw
import os
import webbrowser
import time
import tkinter as tk
from tkinter import messagebox
import subprocess

# --- CONFIGURAZIONE ---
NOME_FOGLIO_PRINCIPALE = "PRONO&RISULTATI"
NOME_FOGLIO_PRONO = "PRONO ORIZZ"
NOME_FOGLIO_COPPA = "CLASSIFICHE COPPA"
CARTELLA_FOTO_NOME = "foto_tabellini"
NOME_FILE_WEB = "visualizza_tabellini.html"

# --- [PARAMETRI ALTEZZA] ---
ALTEZZA_RIFERIMENTO_COPPA = "480px"   
ALTEZZA_ALRE_FOTO = "920px"    

def scatta_foto_precise():
    cartella_lavoro = r"C:\Users\semmd\Documents\GitHub\torneo-coppa-SHECTOR"
    
    if not os.path.exists(cartella_lavoro): os.makedirs(cartella_lavoro)
    os.chdir(cartella_lavoro) 
    
    percorso_foto_sottocartella = os.path.join(cartella_lavoro, CARTELLA_FOTO_NOME)
    percorso_web = os.path.join(cartella_lavoro, NOME_FILE_WEB)
    
    if not os.path.exists(percorso_foto_sottocartella): os.makedirs(percorso_foto_sottocartella)

    try:
        try:
            wb = xw.Book.caller()
        except:
            wb = xw.Book("TURNOINCORSO.xlsm")
            
        sht_main = wb.sheets[NOME_FOGLIO_PRINCIPALE]
        sht_prono = wb.sheets[NOME_FOGLIO_PRONO]
        sht_coppa_foglio = wb.sheets[NOME_FOGLIO_COPPA]
        
        timestamp_url = int(time.time())
        timestamp_label = time.strftime('%d/%m/%Y %H:%M:%S')

        def salva_foto_sicura(foglio, range_celle, cartella_destinazione, nome_file):
            foglio.activate()
            time.sleep(0.7)
            percorso_finale = os.path.join(cartella_destinazione, nome_file)
            foglio.range(range_celle).to_png(percorso_finale)

        # --- SCATTI STANDARD ---
        salva_foto_sicura(sht_main, "B4:T19", percorso_foto_sottocartella, "risultati_partite.png")
        salva_foto_sicura(sht_main, "B24:N38", percorso_foto_sottocartella, "foto_extra.png")
        salva_foto_sicura(sht_main, "AC3:BJ62", percorso_foto_sottocartella, "classifica_turno.png")
        salva_foto_sicura(sht_prono, "A1:AU58", percorso_foto_sottocartella, "I PRONOSTICI DI TUTTI.png")
        
        # --- SCATTI COPPA ---
        # Foto 1a dal foglio PRONO&RISULTATI
        salva_foto_sicura(sht_main, "AD8:AE62", cartella_lavoro, "foto_coppa_1a.png")
        
        # Foto 1b dal foglio PRONO&RISULTATI (Coordinate aggiornate a W4:AA62)
        salva_foto_sicura(sht_main, "W4:AA62", cartella_lavoro, "foto_coppa_1b.png")
        
        # Foto 3 dal foglio CLASSIFICHE COPPA
        salva_foto_sicura(sht_coppa_foglio, "L175:AA220", cartella_lavoro, "foto_coppa_3.png")

        # --- TABELLINI ---
        lista_nomi = sht_main.range("AD10:AD62").value
        riga_5 = sht_main.range("5:5").value
        nomi_processati = ["I PRONOSTICI DI TUTTI"]
        
        for nome in lista_nomi:
            if nome:
                try:
                    nome_f = str(nome).strip()
                    col_base = riga_5.index(nome) + 1
                    range_tab = sht_main.range((5, col_base - 4), (26, col_base + 6))
                    range_tab.to_png(os.path.join(percorso_foto_sottocartella, f"{nome_f}.png"))
                    nomi_processati.append(nome_f)
                except: continue

        menu_options = "".join([f'<option value="{CARTELLA_FOTO_NOME}/{n}.png?v={timestamp_url}">{n}</option>' for n in nomi_processati])
        
        # --- HTML ---
        html_content = f'''
        <!DOCTYPE html>
        <html lang="it">
        <head>
            <meta charset="UTF-8">
            <title>16° Torneo Pronostici - 2025/26</title>
            <style>
                body {{ background-color: #121212; color: white; font-family: "AR CENA", sans-serif; text-align: center; margin: 0; padding: 15px; }}
                .settore-verde {{ border: 3px solid #00ff88; background: #0a0a0a; padding: 30px; border-radius: 15px; margin-bottom: 40px; }}
                .flex-top {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 30px; max-width: 1400px; margin: 0 auto; }}
                .img-risultati {{ width: 100%; max-width: 850px; border: 2px solid #00ff88; border-radius: 8px; }}
                .img-extra {{ width: 100%; max-width: 450px; border: 1px solid #444; border-radius: 8px; }}
                .img-max {{ width: 100%; border: 2px solid #00ff88; border-radius: 5px; }}
                select {{ padding: 15px; border-radius: 8px; background: #1a1a1a; color: #00ff88; border: 2px solid #00ff88; font-size: 22px; width: 90%; max-width: 600px; }}
            </style>
        </head>
        <body>
            <div class="settore-verde">
                <h1 style="color: #00ff88; font-size: 52px; text-shadow: 3px 3px #004422;">🏆 16° Torneo PRONOSTICI SERIE A - 2025/26 🏆</h1>
            </div>

            <div class="settore-verde">
                <span style="color: #00ff88; font-size: 32px; text-decoration: underline;">⚽ I RISULTATI ⚽</span>
                <div class="flex-top" style="margin-top:20px;">
                    <div style="flex: 2;"><img src="{CARTELLA_FOTO_NOME}/risultati_partite.png?v={timestamp_url}" class="img-risultati"></div>
                    <div style="flex: 1;"><img src="{CARTELLA_FOTO_NOME}/foto_extra.png?v={timestamp_url}" class="img-extra"></div>
                </div>
            </div>

            <div class="settore-verde">
                <span style="color: #00ff88; font-size: 32px; text-decoration: underline;">📊 LE CLASSIFICHE 📊</span>
                <img src="{CARTELLA_FOTO_NOME}/classifica_turno.png?v={timestamp_url}" class="img-max" style="margin-top:20px;">
            </div>

            <div class="settore-verde">
                <span style="color: #00ff88; font-size: 32px; text-decoration: underline;">📝 I TABELLINI 📝</span><br><br>
                <select onchange="var img=document.getElementById('f-tab'); img.src=this.value; img.style.display='inline-block';">
                    <option value="">-- SELEZIONA GIOCATORE --</option>
                    {menu_options}
                </select>
                <br><img id="f-tab" style="display:none; max-width:850px; width:95%; border: 2px solid #00ff88; margin-top:25px; border-radius:12px;">
            </div>

            <hr style="border: 5px solid gold; margin: 60px 0;">
            
            <div style="border: 3px solid gold; background: #0a0a0a; padding: 30px; border-radius: 15px;">
                <h1 style="color: gold; font-size: 52px;">🏆 1° Coppa SHECTOR 🏆</h1>
                
                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold; margin-bottom: 40px;">
                    <div style="display: flex; justify-content: center; align-items: flex-start; gap: 20px;">
                        <img src="immagini_fisse/Primafoto.png" style="height: {ALTEZZA_RIFERIMENTO_COPPA}; border: 2px solid gold; border-radius: 5px;">
                        <img src="foto_coppa_1a.png?v={timestamp_url}" style="height: {ALTEZZA_ALRE_FOTO}; border: 1px solid #444; border-radius: 5px;">
                        <img src="foto_coppa_1b.png?v={timestamp_url}" style="height: {ALTEZZA_ALRE_FOTO}; border: 1px solid #444; border-radius: 5px;">
                    </div>
                </div>

                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold; margin-bottom: 40px;">
                    <div style="display: flex; justify-content: center; gap: 20px;">
                        <div style="width: 28%;">
                            <img src="immagini_fisse/2afase.png" style="width: 100%;">
                            <img src="immagini_fisse/CLASS2AFASE.png" style="width: 80%; border: 1px solid gold;">
                        </div>
                        <img src="foto_coppa_2.png?v={timestamp_url}" style="width: 50%; border: 2px solid #444;">
                    </div>
                </div>

                <div style="background: #1a1a00; padding: 25px; border-radius: 15px; border: 1px solid gold;">
                    <div style="display: flex; justify-content: center; align-items: flex-start; gap: 20px;">
                        <img src="immagini_fisse/3afase.png" style="height: {ALTEZZA_RIFERIMENTO_COPPA}; border: 2px solid gold; border-radius: 5px;">
                        <img src="foto_coppa_3.png?v={timestamp_url}" style="width: 65%; border: 2px solid gold; border-radius: 5px;">
                    </div>
                </div>
            </div>
            <p style="color:#00ff88; margin-top:20px;">Ultimo aggiornamento: {timestamp_label}</p>
        </body>
        </html>
        '''
        
        with open(percorso_web, "w", encoding="utf-8") as f: f.write(html_content)
        webbrowser.open(f"file:///{percorso_web.replace(os.sep, '/')}")

        # --- GITHUB ---
        p_git = r"C:\Program Files\Git\bin\git.exe"
        git_cmd = p_git if os.path.exists(p_git) else "git"
        subprocess.run([git_cmd, "add", "--all"], cwd=cartella_lavoro)
        subprocess.run([git_cmd, "commit", "-m", f"Update {timestamp_label}"], cwd=cartella_lavoro)
        subprocess.run([git_cmd, "pull", "--rebase", "origin", "main"], cwd=cartella_lavoro)
        subprocess.run([git_cmd, "push", "origin", "main"], cwd=cartella_lavoro)
        
    except Exception as e:
        root = tk.Tk(); root.withdraw(); messagebox.showerror("ERRORE", str(e)); root.destroy()

if __name__ == "__main__":
    scatta_foto_precise()