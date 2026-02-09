import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar
import pandas as pd
import random
import unicodedata
import threading
import queue
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options as FirefoxOptions

# Intentar importar Pillow para capturas
try:
    from PIL import ImageGrab
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

# --- ESTILOS VISUALES ---
STYLE = {
    "bg_app": "#121212",        
    "bg_panel": "#1e1e1e",      
    "bg_header": "#2d2d30",
    "bg_cell_empty": "#1f1f1f",       
    "fg_prim": "#ffffff",       
    "fg_sec": "#aaaaaa",        
    "accent": "#bb86fc",        
    "avail": "#03dac6",         
    "danger": "#cf6679",        
    "warning": "#ffb74d",       
    "success": "#81c784",       
    "info": "#64b5f6",
    "switch": "#ff9800",
    "font_h": ("Segoe UI", 10, "bold"),
    "font_txt": ("Segoe UI", 9),
    "font_small": ("Segoe UI", 8),
    "font_badge": ("Arial", 7, "bold")
}

class CollapsiblePane(tk.Frame):
    def __init__(self, parent, title, expanded=True):
        super().__init__(parent, bg=STYLE["bg_panel"])
        self.expanded = expanded
        self.header_frame = tk.Frame(self, bg=STYLE["bg_header"], pady=2, padx=5, relief="raised", bd=1)
        self.header_frame.pack(fill="x", side="top")
        self.btn_toggle = tk.Button(self.header_frame, text="[ - ]" if expanded else "[ + ]", 
                                    command=self.toggle, bg=STYLE["bg_header"], fg=STYLE["accent"], 
                                    relief="flat", font=("Consolas", 9, "bold"), width=5)
        self.btn_toggle.pack(side="left")
        self.lbl_title = tk.Label(self.header_frame, text=title, bg=STYLE["bg_header"], fg="white", font=STYLE["font_h"])
        self.lbl_title.pack(side="left", padx=5)
        self.content_frame = tk.Frame(self, bg=STYLE["bg_panel"])
        if expanded: self.content_frame.pack(fill="both", expand=True)

    def toggle(self):
        if self.expanded:
            self.content_frame.pack_forget()
            self.btn_toggle.config(text="[ + ]")
            self.pack(fill="x", expand=False)
        else:
            self.content_frame.pack(fill="both", expand=True)
            self.btn_toggle.config(text="[ - ]")
            self.pack(fill="both", expand=True)
        self.expanded = not self.expanded

class LoginWindow(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.parent = parent
        self.callback = callback
        self.result = None
        
        self.title("Login ESPOL")
        self.geometry("500x600")
        self.configure(bg=STYLE["bg_app"])
        self.resizable(False, False)
        
        # Centrar ventana
        self.transient(parent)
        self.grab_set()
        
        self._create_widgets()
        
    def _create_widgets(self):
        # Header
        header_frame = tk.Frame(self, bg=STYLE["accent"], height=80)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔐 Login ESPOL", bg=STYLE["accent"], 
                fg="white", font=("Segoe UI", 16, "bold")).pack(expand=True, pady=(10,0))
        tk.Label(header_frame, text="Ingresa tus credenciales del portal académico", 
                bg=STYLE["accent"], fg="white", font=("Segoe UI", 10)).pack(expand=True, pady=(0,10))
        
        # Main content
        content_frame = tk.Frame(self, bg=STYLE["bg_app"])
        content_frame.pack(fill="both", expand=True, padx=30, pady=15)
        
        # Usuario
        tk.Label(content_frame, text="Usuario ESPOL:", bg=STYLE["bg_app"], 
                fg=STYLE["fg_prim"], font=STYLE["font_h"], anchor="w").pack(fill="x", pady=(0,5))
        self.user_var = StringVar()
        self.user_entry = tk.Entry(content_frame, textvariable=self.user_var, 
                                  font=("Segoe UI", 11), bg="#2a2a2a", fg="white",
                                  insertbackground="white", relief="flat", highlightthickness=1,
                                  highlightcolor=STYLE["accent"], highlightbackground="#444")
        self.user_entry.pack(fill="x", pady=(0,15), ipady=8)
        
        # Contraseña
        tk.Label(content_frame, text="Contraseña:", bg=STYLE["bg_app"], 
                fg=STYLE["fg_prim"], font=STYLE["font_h"], anchor="w").pack(fill="x", pady=(0,5))
        self.pass_var = StringVar()
        self.pass_entry = tk.Entry(content_frame, textvariable=self.pass_var, 
                                  font=("Segoe UI", 11), bg="#2a2a2a", fg="white",
                                  insertbackground="white", show="•", relief="flat", 
                                  highlightthickness=1, highlightcolor=STYLE["accent"], 
                                  highlightbackground="#444")
        self.pass_entry.pack(fill="x", pady=(0,15), ipady=8)
        
        # Año
        tk.Label(content_frame, text="Año Académico:", bg=STYLE["bg_app"], 
                fg=STYLE["fg_prim"], font=STYLE["font_h"], anchor="w").pack(fill="x", pady=(0,5))
        self.year_var = StringVar(value=str(datetime.now().year))
        self.year_combo = ttk.Combobox(content_frame, textvariable=self.year_var, 
                                      values=[str(y) for y in range(2020, 2030)], 
                                      state="readonly", font=("Segoe UI", 11))
        self.year_combo.pack(fill="x", pady=(0,15))
        
        # Término
        tk.Label(content_frame, text="Término Académico:", bg=STYLE["bg_app"], 
                fg=STYLE["fg_prim"], font=STYLE["font_h"], anchor="w").pack(fill="x", pady=(0,5))
        self.term_var = StringVar(value="Término I")
        self.term_combo = ttk.Combobox(content_frame, textvariable=self.term_var, 
                                      values=["Término I", "Término II", "Término Vacacional"], 
                                      state="readonly", font=("Segoe UI", 11))
        self.term_combo.pack(fill="x", pady=(0,15))
        
        # Materias
        tk.Label(content_frame, text="Códigos de Materias (separados por comas):", 
                bg=STYLE["bg_app"], fg=STYLE["fg_prim"], font=STYLE["font_h"], 
                anchor="w").pack(fill="x", pady=(0,5))
        self.subjects_var = StringVar()
        self.subjects_entry = tk.Entry(content_frame, textvariable=self.subjects_var, 
                                      font=("Segoe UI", 11), bg="#2a2a2a", fg="white",
                                      insertbackground="white", relief="flat", 
                                      highlightthickness=1, highlightcolor=STYLE["accent"], 
                                      highlightbackground="#444")
        self.subjects_entry.pack(fill="x", pady=(0,20), ipady=8)
        self.subjects_entry.insert(0, "CCPG1053,SOFG1009,CCPG1055,CCPG1041,CCPG1044,CCPG1048") #6 Materias prellenadas
        
        # Footer info
        tk.Label(content_frame, text="⚠ Nota: El proceso puede tomar varios minutos", 
                bg=STYLE["bg_app"], fg=STYLE["warning"], font=("Segoe UI", 9)).pack(pady=(10,0))
        
        # Botones FRAME - Ahora en la parte inferior de la ventana
        btn_frame = tk.Frame(self, bg=STYLE["bg_app"], pady=15, padx=30)
        btn_frame.pack(fill="x", side="bottom", pady=10)
        
        tk.Button(btn_frame, text="Cancelar", bg="#444", fg="white",
                 font=("Segoe UI", 10, "bold"), relief="flat", padx=20,
                 command=self.destroy).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="Iniciar Scraping", bg=STYLE["accent"], 
                 fg="white", font=("Segoe UI", 10, "bold"), relief="flat", 
                 padx=20, command=self.on_login).pack(side="right", padx=5)
        
    def on_login(self):
        user = self.user_var.get().strip()
        password = self.pass_var.get().strip()
        year = self.year_var.get().strip()
        term = self.term_var.get()
        subjects = [s.strip() for s in self.subjects_var.get().split(",") if s.strip()]
        
        if not user or not password:
            messagebox.showerror("Error", "Usuario y contraseña son obligatorios")
            return
            
        if not subjects:
            messagebox.showerror("Error", "Debe ingresar al menos una materia")
            return
            
        # Mapear término a ID
        term_map = {
            "Término I": "ctl00_contenido_listab_1",
            "Término II": "ctl00_contenido_listab_2",
            "Término Vacacional": "ctl00_contenido_listab_0"
        }
        
        self.result = {
            "user": user,
            "password": password,
            "year": year,
            "term_id": term_map.get(term, "ctl00_contenido_listab_1"),
            "term_name": term,
            "subjects": subjects
        }
        
        self.destroy()
        self.callback(self.result)

class CaptchaViewerWindow(tk.Toplevel):
    def __init__(self, parent, driver):
        super().__init__(parent)
        self.parent = parent
        self.driver = driver
        self.solved = False
        
        self.title("CAPTCHA ESPOL")
        self.geometry("500x400")
        self.configure(bg=STYLE["bg_app"])
        
        # Centrar ventana
        self.transient(parent)
        self.grab_set()
        
        # Traer el navegador al frente automáticamente
        self.bring_browser_to_front()
        
        self._create_widgets()
        
    def _create_widgets(self):
        # Header
        header_frame = tk.Frame(self, bg=STYLE["warning"], height=60)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔍 CAPTCHA ESPOL", bg=STYLE["warning"], 
                fg="black", font=("Segoe UI", 14, "bold")).pack(expand=True, pady=10)
        
        # Contenedor principal
        main_frame = tk.Frame(self, bg=STYLE["bg_app"], padx=20, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # Instructions
        tk.Label(main_frame, text="El navegador se ha abierto para que resuelvas el CAPTCHA.", 
                bg=STYLE["bg_app"], fg="white", font=("Segoe UI", 10)).pack(anchor="w", pady=(0,10))
        
        tk.Label(main_frame, text="Instrucciones:", 
                bg=STYLE["bg_app"], fg=STYLE["avail"], font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,5))
        
        instructions = [
            "1. Resuelve el CAPTCHA en el navegador",
            "2. Haz clic en 'Iniciar sesión' en el navegador",
            "3. Espera a que cargue la página principal",
            "4. Luego haz clic en 'Continuar' aquí"
        ]
        
        for instruction in instructions:
            tk.Label(main_frame, text=instruction, 
                    bg=STYLE["bg_app"], fg=STYLE["fg_sec"], 
                    font=("Segoe UI", 9), anchor="w").pack(anchor="w", pady=2)
        
        # Nota importante
        tk.Label(main_frame, text="⚠ NO CIERRES EL NAVEGADOR", 
                bg=STYLE["bg_app"], fg=STYLE["warning"], 
                font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(10,5))
        
        tk.Label(main_frame, text="El navegador debe permanecer abierto para continuar.", 
                bg=STYLE["bg_app"], fg=STYLE["warning"], font=("Segoe UI", 9)).pack(anchor="w", pady=(0,20))
        
        # Botón Continuar - CENTRADO y VISIBLE
        btn_container = tk.Frame(main_frame, bg=STYLE["bg_app"], pady=20)
        btn_container.pack(side="bottom", fill="x")
        
        self.continue_btn = tk.Button(btn_container, text="CONTINUAR", bg=STYLE["success"], 
                 fg="white", font=("Segoe UI", 12, "bold"), relief="flat", 
                 padx=40, pady=12, command=self.on_continue)
        self.continue_btn.pack()
        
    def bring_browser_to_front(self):
        """Trae el navegador al frente"""
        try:
            # Traer el navegador al frente
            self.driver.switch_to.window(self.driver.current_window_handle)
            self.driver.maximize_window()
        except Exception as e:
            print(f"Error al traer navegador al frente: {e}")
        
    def on_continue(self):
        self.solved = True
        self.destroy()

class FastScraperThread(threading.Thread):
    def __init__(self, login_data, progress_callback, captcha_callback, finish_callback):
        super().__init__()
        self.login_data = login_data
        self.progress_callback = progress_callback
        self.captcha_callback = captcha_callback
        self.finish_callback = finish_callback
        self.driver = None
        self.running = True
        self.output_file = "INFORME.csv"
        
    def run(self):
        try:
            self._execute_fast_scraper()
        except Exception as e:
            self.progress_callback(f"❌ Error: {str(e)}", "error")
            self.finish_callback(False, str(e))
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
    
    def _execute_fast_scraper(self):
        # Configurar driver
        self.progress_callback("🚀 Iniciando navegador...", "info")
        start_time = time.time()
        
        # Configuración del navegador
        options = FirefoxOptions()
        options.add_argument("--start-maximized")
        self.driver = webdriver.Firefox(options=options)
        self.driver.set_page_load_timeout(30)
        
        # ---------- DECLARACIÓN DE VARIABLES -------------------
        link = "https://www.academico.espol.edu.ec/login.aspx?ReturnUrl=%2fUI%2fInformacionAcademica%2finformaciongeneral.aspx"
        
        self.progress_callback("🌐 Accediendo al portal académico...", "info")
        self.driver.get(link)
        wait = WebDriverWait(self.driver, 20)
        
        # ---------- INTERFAZ DE INICIO -----------------
        # Ingresar usuario
        self.progress_callback("🔑 Ingresando credenciales...", "info")
        caja_busqueda = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$contenido$txtuser")))
        caja_busqueda.send_keys(self.login_data["user"])
        
        # Clickear siguiente
        self.driver.find_element(By.NAME, "ctl00$contenido$btnSigte").click()
        
        # Ingresar contraseña
        caja_access = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$contenido$txtpsw")))
        caja_access.send_keys(self.login_data["password"])
        
        # Mostrar ventana para que usuario resuelva el CAPTCHA
        self.progress_callback("👤 Por favor, resuelva el CAPTCHA en el navegador...", "warning")
        self.captcha_callback("manual", self.driver)
        
        # El usuario ya hizo clic en "Iniciar sesión" en el navegador
        self.progress_callback("⏳ Esperando que se complete el inicio de sesión...", "info")
        
        # Esperar a que se cargue la página principal
        try:
            wait.until(lambda driver: driver.current_url != link)
            current_url = self.driver.current_url
            if "login.aspx" in current_url:
                self.progress_callback("❌ Error: No se pudo iniciar sesión después del CAPTCHA", "error")
                raise Exception("No se pudo iniciar sesión después del CAPTCHA")
            else:
                self.progress_callback("✅ Inicio de sesión exitoso", "success")
        except TimeoutException:
            self.progress_callback("⚠ Tiempo de espera agotado, intentando continuar...", "warning")
        
        # ---------- DECLARACIÓN DE FUNCIONES ----------------------------
        def schedule(filas_horario):
            """Recibe una tabla e itera x filas y columnas. Retorna una lista de cada fila con los valores de las columnas separadas por comas"""
            resultado = []
            for fila in filas_horario:
                valores = []
                for col in fila.find_elements(By.TAG_NAME, "td"):  # columnas dentro de la fila
                    col_text = col.text
                    if "CAMPUS" in col_text:
                        col_text = col_text.split(" CAMPUS")[0]
                        valores.append(col_text)
                    else:
                        valores.append(col_text)
                resultado.append(",".join(valores))
            return resultado
            
        def existe_boton(driver, id):
            """Recibe el driver y un ID. Retorna falso si el ID no existe"""
            try:
                driver.find_element(By.ID, id)
                return True
            except NoSuchElementException:
                return False
        
        # ---------- ITERAR POR CÓDIGO DE MATERIAS ------------------
        self.progress_callback(f"📚 Procesando {len(self.login_data['subjects'])} materias...", "info")
        
        with open(self.output_file, "w", newline="", encoding="UTF-8") as file:
            # Encabezado
            file.write("Materia,Docente,Paralelo,Modalidad,Dia_T1,HoraI_T1,HoraF_T1,Aula_T1,Edificio_T1,Dia_T2,HoraI_T2,HoraF_T2,Aula_T2,Edificio_T2,Dia_P,HoraI_P,HoraF_P,Aula_P,Edificio_P,Examen1,Aula1,Examen2,Aula2,Examen3,Aula3\n")
            
            for i, signature_code in enumerate(self.login_data["subjects"]):
                if not self.running:
                    break
                    
                self.progress_callback(f"📖 Materia {i+1}/{len(self.login_data['subjects'])}: {signature_code}", "info")
                
                try:
                    # Localiza el enlace por el texto visible "REGISTROS"
                    registro_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "REGISTROS")))
                    registro_link.click()

                    # Ingresar año
                    caja_year = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$contenido$txtAnio")))
                    caja_year.clear()
                    caja_year.send_keys(self.login_data["year"])

                    # Clickear término
                    self.driver.find_element(By.ID, self.login_data["term_id"]).click()

                    # Clickear busquedad
                    self.driver.find_element(By.NAME, "ctl00$contenido$btnConsultar").click()

                    # Clickear Busquedad por Código
                    self.driver.find_element(By.ID, "ctl00_contenido_RBList_1").click()

                    # Enviar codigo materia
                    code = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$contenido$codigoMateria")))
                    code.clear()
                    code.send_keys(signature_code)

                    # Clickear Buscar
                    self.driver.find_element(By.NAME, "ctl00$contenido$Button2").click()

                    # Encuentra todos los enlaces de paralelos dentro de esa celda
                    try:
                        paralelos = wait.until(EC.presence_of_all_elements_located(
                            (By.XPATH, "//table[@id='ctl00_contenido_tbHorarios']/tbody/tr/td[3]/a")))
                    except:
                        self.progress_callback(f"⚠ No se encontraron paralelos para {signature_code}", "warning")
                    
                    self.progress_callback(f"  📊 Encontrados {len(paralelos)} paralelos", "info")

                    # Iterar sobre cada paralelo
                    for j in range(len(paralelos)):
                        if not self.running:
                            break
                            
                        try:
                            par_Prac = str(100 + j + 1)
                            # Re-localizar paralelos y hacer click
                            paralelos_links = self.driver.find_elements(By.XPATH, 
                                "//table[@id='ctl00_contenido_tbHorarios']/tbody/tr/td[3]/a")
                            if j >= len(paralelos_links):
                                break
                                
                            paralelo = paralelos_links[j]
                            paralelo.click()
                            
                            par_Teo_display = int(self.driver.find_element(By.ID, "ctl00_contenido_LabelParalelo").text)
                            
                            # ---SCRAPEAR---
                            docente = self.driver.find_element(By.ID, "ctl00_contenido_LabelProfesor").text
                            materia = self.driver.find_element(By.ID, "ctl00_contenido_LabelNombreMateria").text
                            
                            # MOSTRAR MENSAJE CON NOMBRE DE MATERIA
                            if par_Teo_display == j + 1:
                                self.progress_callback(f"  ➡️ Entrando a {materia[:20]}... paralelo {j+1}", "info")
                            
                            modalidad = self.driver.find_element(By.ID, "ctl00_contenido_modo_curso").text.split("MODALIDAD ")[1]
                            
                            # EXAMENES
                            exm1, aula_1 = "", ""
                            exm2, aula_2 = "", ""
                            exm3, aula_3 = "", ""
                            
                            try:
                                exm1 = self.driver.find_element(By.ID, "ctl00_contenido_LabelParcial").text
                                aula_1 = self.driver.find_element(By.ID, "ctl00_contenido_aulaParcial").text.split(" CAMPUS")[0]
                            except:
                                pass
                                
                            try:
                                exm2 = self.driver.find_element(By.ID, "ctl00_contenido_LabelFinal").text
                                aula_2 = self.driver.find_element(By.ID, "ctl00_contenido_aulaFinal").text.split(" CAMPUS")[0]
                            except:
                                pass
                                
                            try:
                                exm3 = self.driver.find_element(By.ID, "ctl00_contenido_LabelMejora").text
                                aula_3 = self.driver.find_element(By.ID, "ctl00_contenido_aulaMejora").text.split(" CAMPUS")[0]
                            except:
                                pass
                            
                            # Horarios del teórico
                            horariosT = self.driver.find_elements(By.XPATH, "//table[@id='ctl00_contenido_TableHorarios']/tbody")
                            
                            # Abrir horarios práctico
                            horariosP = []
                            if existe_boton(self.driver, par_Prac):
                                self.driver.find_element(By.ID, par_Prac).click()
                                id_table = f"tabla_{par_Prac}"
                                horariosP = self.driver.find_elements(By.XPATH, f"//div[@id='{id_table}']/table[@class='display']/tbody")
                            
                            # Procesar horarios
                            for teo in schedule(horariosT):  # teo <- filas & prac <- columnas
                                count_T = int(len(teo.split(",")) / 5)  # Cantidad de horarios teóricos
                                
                                if existe_boton(self.driver, par_Prac):
                                    for prac in schedule(horariosP):
                                        if count_T == 2:
                                            file.write(f"{materia},{docente},{par_Teo_display},{modalidad},{teo},{prac},{exm1},{aula_1},{exm2},{aula_2},{exm3},{aula_3}\n")
                                        elif count_T == 1:
                                            file.write(f"{materia},{docente},{par_Teo_display},{modalidad},{teo},,,,,,{prac},{exm1},{aula_1},{exm2},{aula_2},{exm3},{aula_3}\n")
                                else:
                                    if count_T == 2:
                                        file.write(f"{materia},{docente},{par_Teo_display},{modalidad},{teo},,,,,,{exm1},{aula_1},{exm2},{aula_2},{exm3},{aula_3}\n")
                                    elif count_T == 1 and not existe_boton(self.driver, par_Prac):
                                        file.write(f"{materia},{docente},{par_Teo_display},{modalidad},{teo},,,,,,,,,,,{exm1},{aula_1},{exm2},{aula_2},{exm3},{aula_3}\n")
                            
                        except Exception as e:
                            self.progress_callback(f"  ⚠ Error procesando paralelo {j+1}: {str(e)}", "warning")
                        
                        # Volver atrás
                        self.driver.back()
                        
                except Exception as e:
                    self.progress_callback(f"⚠ Error con materia {signature_code}: {str(e)}", "warning")
                    continue
        
        # Cerrar sesión
        try:
            self.progress_callback("👋 Cerrando sesión...", "info")
            logout_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_lbtSalir")))
            logout_btn.click()
            
        except:
            self.progress_callback("⚠ No se pudo cerrar sesión automáticamente", "warning")
        
        # Cerrar navegador
        self.driver.quit()
        
        
        elapsed_time = time.time() - start_time
        self.progress_callback(f"✅ Scraping completado en {elapsed_time:.1f} segundos", "success")
        self.finish_callback(True, self.output_file)
        
    def stop(self):
        self.running = False
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass

class AcademicScheduler:
    def __init__(self, root):
        self.root = root
        self.root.title("Planificador Académico ESPOL")
        try: 
            self.root.state('zoomed')
        except: 
            self.root.attributes('-fullscreen', True) 
        self.root.configure(bg=STYLE["bg_app"])
        
        # Variables para el scraper
        self.scraper_thread = None
        self.captcha_window = None
        self.captcha_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        
        # Datos
        self.df = pd.DataFrame()
        self.df_view = pd.DataFrame()
        self.inscritas = set()
        self.materias_inscritas = set()
        self.color_map = {}
        self.initials_map = {}  # Mapa para iniciales únicas
        
        # Filtros
        self.f_materia = tk.StringVar()
        self.f_docente = tk.StringVar()
        self.f_dia = tk.StringVar()
        self.count_var = tk.StringVar(value="Materias Registradas: 0")
        
        # Sugerencias
        self.sug_turno_var = tk.StringVar()
        self.sug_docente_var = tk.StringVar()
        
        self._init_ui()
        self._start_queue_processor()
    
    def _init_ui(self):
        # HEADER
        header = tk.Frame(self.root, bg=STYLE["bg_panel"], pady=8, padx=15)
        header.pack(fill="x")
        btn_frame = tk.Frame(header, bg=STYLE["bg_panel"])
        btn_frame.pack(side="left")
        
        self._btn(btn_frame, "🔍 Scraping Académico", self.open_scraper_dialog, STYLE["accent"], "#000")
        self._btn(btn_frame, "📂 Cargar CSV", self.load_csv, STYLE["info"], "#000")
        
        if HAS_PILLOW: 
            self._btn(btn_frame, "📷 Capturar", self.save_screenshot, "#333", "white")
        
        self._btn(btn_frame, "📅 Exámenes", self.open_exam_window, "#333", STYLE["avail"])
        
        # Botón de Cerrar - MEJORADO para cancelar todo
        self._btn(btn_frame, "🚫 Cerrar Todo", self.force_quit_app, STYLE["danger"], "white")
        
        # Filtros
        self._crear_filtro_dinamico(header, "Materia", self.f_materia, 20)
        self._crear_filtro_dinamico(header, "Docente", self.f_docente, 20)
        self._crear_filtro_dinamico(header, "Día", self.f_dia, 12)
        
        self._btn(header, "↺ Reset", self.reset_filters, "#444", "white")
        self._btn(header, "🗑 Limpiar", self.clear_schedule, STYLE["danger"], "white")
        
        tk.Label(header, textvariable=self.count_var, bg="#2d2d30", fg=STYLE["avail"], 
                font=("Segoe UI", 11, "bold"), padx=15, pady=5).pack(side="right", padx=10)
        
        # Progress Frame (oculto inicialmente)
        self.progress_frame = tk.Frame(self.root, bg=STYLE["bg_panel"], height=70)
        self.progress_label = tk.Label(self.progress_frame, text="", bg=STYLE["bg_panel"], 
                                      fg="white", font=("Segoe UI", 10, "bold"), pady=5)
        self.progress_label.pack()
        
        # Control buttons frame
        self.control_frame = tk.Frame(self.progress_frame, bg=STYLE["bg_panel"])
        
        # BODY
        self.body = tk.PanedWindow(self.root, bg=STYLE["bg_app"], orient="horizontal", 
                                  sashwidth=4, sashrelief="flat")
        self.body.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Grilla
        self.grid_container = tk.Frame(self.body, bg=STYLE["bg_app"])
        self.body.add(self.grid_container, stretch="always")
        self.cell_frames = {}
        self._build_grid_structure()
        
        # Panel Lateral
        self.side_panel = tk.Frame(self.body, bg=STYLE["bg_panel"], width=480)
        self.body.add(self.side_panel, stretch="never")
        self.side_panel.pack_propagate(False)
        
        # SECCIÓN 1: DETALLES
        self.pane_detalles = CollapsiblePane(self.side_panel, "DETALLES Y OPCIONES")
        self.pane_detalles.pack(fill="both", expand=True, pady=2)
        self.canvas_det = tk.Canvas(self.pane_detalles.content_frame, bg=STYLE["bg_panel"], 
                                   highlightthickness=0)
        scroll_det = ttk.Scrollbar(self.pane_detalles.content_frame, orient="vertical", 
                                  command=self.canvas_det.yview)
        self.frame_list = tk.Frame(self.canvas_det, bg=STYLE["bg_panel"])
        self.frame_list.bind("<Configure>", lambda e: self.canvas_det.configure(
            scrollregion=self.canvas_det.bbox("all")))
        self.canvas_det.create_window((0, 0), window=self.frame_list, anchor="nw", width=440)
        self.canvas_det.configure(yscrollcommand=scroll_det.set)
        self.canvas_det.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_det.pack(side="right", fill="y")
        self.lbl_placeholder = tk.Label(self.frame_list, text="Selecciona una casilla...", 
                                       bg=STYLE["bg_panel"], fg="#666", pady=20)
        self.lbl_placeholder.pack()
        
        # SECCIÓN 2: SUGERENCIAS
        self.pane_sug = CollapsiblePane(self.side_panel, "✨ SUGERENCIAS COMPATIBLES")
        self.pane_sug.pack(fill="both", expand=True, pady=2)
        
        filter_frame = tk.Frame(self.pane_sug.content_frame, bg=STYLE["bg_panel"], pady=5)
        filter_frame.pack(fill="x", padx=5)
        tk.Label(filter_frame, text="Turno:", bg=STYLE["bg_panel"], fg="#aaa", 
                font=STYLE["font_small"]).pack(side="left", padx=2)
        cb_turno = ttk.Combobox(filter_frame, textvariable=self.sug_turno_var, 
                               values=["Cualquiera", "Mañana (<12)", "Tarde (12-17)", "Noche (>17)"], 
                               state="readonly", width=12)
        cb_turno.current(0)
        cb_turno.pack(side="left", padx=5)
        cb_turno.bind("<<ComboboxSelected>>", lambda e: self.generate_suggestions())
        tk.Label(filter_frame, text="Docente:", bg=STYLE["bg_panel"], fg="#aaa", 
                font=STYLE["font_small"]).pack(side="left", padx=2)
        self.cb_sug_doc = ttk.Combobox(filter_frame, textvariable=self.sug_docente_var, width=15)
        self.cb_sug_doc.pack(side="left", padx=5)
        self.cb_sug_doc.bind('<KeyRelease>', lambda e: self.on_sug_doc_key(e))
        self.cb_sug_doc.bind("<<ComboboxSelected>>", lambda e: self.generate_suggestions())
        
        self.canvas_sug = tk.Canvas(self.pane_sug.content_frame, bg=STYLE["bg_panel"], 
                                   highlightthickness=0)
        scroll_sug = ttk.Scrollbar(self.pane_sug.content_frame, orient="vertical", 
                                  command=self.canvas_sug.yview)
        self.frame_sug = tk.Frame(self.canvas_sug, bg=STYLE["bg_panel"])
        self.frame_sug.bind("<Configure>", lambda e: self.canvas_sug.configure(
            scrollregion=self.canvas_sug.bbox("all")))
        self.canvas_sug.create_window((0, 0), window=self.frame_sug, anchor="nw", width=440)
        self.canvas_sug.configure(yscrollcommand=scroll_sug.set)
        self.canvas_sug.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_sug.pack(side="right", fill="y")
        
        # SECCIÓN 3: MONITOR GLOBAL
        self.pane_global = CollapsiblePane(self.side_panel, "⚠ AVISOS")
        self.pane_global.pack(fill="both", expand=True, pady=2)
        frame_txt = tk.Frame(self.pane_global.content_frame, bg=STYLE["bg_panel"])
        frame_txt.pack(fill="both", expand=True, padx=5, pady=5)
        scroll_g = ttk.Scrollbar(frame_txt)
        scroll_g.pack(side="right", fill="y")
        self.txt_global = tk.Text(frame_txt, bg="#1e1e1e", fg="#ccc", font=("Consolas", 8), 
                                 relief="flat", state="disabled", height=12, 
                                 yscrollcommand=scroll_g.set)
        self.txt_global.pack(side="left", fill="both", expand=True)
        scroll_g.config(command=self.txt_global.yview)
    
    def _btn(self, parent, text, cmd, bg, fg):
        tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, 
                 font=("Segoe UI", 9, "bold"), relief="flat", padx=10).pack(side="left", padx=5)
    
    def force_quit_app(self):
        """Cierra la aplicación forzosamente, cancelando todos los procesos"""
        if self.scraper_thread and self.scraper_thread.is_alive():
            if messagebox.askyesno("Confirmar Cierre", 
                                  "⚠ Hay un proceso de scraping en ejecución.\n\n" +
                                  "¿Deseas cancelar TODOS los procesos y cerrar la aplicación?"):
                self.scraper_thread.stop()
                # Esperar un momento para que el thread se detenga
                self.root.destroy()
        else:
            self.root.destroy()
    
    def _crear_filtro_dinamico(self, parent, label, variable, width):
        f = tk.Frame(parent, bg=STYLE["bg_panel"])
        f.pack(side="left", padx=5)
        tk.Label(f, text=label, bg=STYLE["bg_panel"], fg=STYLE["fg_sec"], 
                font=("Segoe UI", 7)).pack(anchor="w")
        cb = ttk.Combobox(f, textvariable=variable, width=width)
        cb.pack()
        safe_name = ''.join(c for c in unicodedata.normalize('NFD', label) if unicodedata.category(c) != 'Mn').lower()
        setattr(self, f"cb_{safe_name}", cb)
        cb.bind('<KeyRelease>', lambda e, var=variable, cbox=cb: self.on_key_release(e, var, cbox))
        cb.bind("<<ComboboxSelected>>", self.on_filter_change)
    
    def open_scraper_dialog(self):
        """Abre la ventana de login para el scraping"""
        if self.scraper_thread and self.scraper_thread.is_alive():
            messagebox.showwarning("Advertencia", "Ya hay un proceso de scraping en ejecución")
            return
            
        login_window = LoginWindow(self.root, self.on_login_data_received)
    
    def on_login_data_received(self, login_data):
        """Callback cuando se reciben datos de login"""
        self.show_progress_frame()
        self.progress_label.config(text="🚀 Iniciando scraping manual...", fg=STYLE["warning"])
        
        # Configurar botones de control
        for widget in self.control_frame.winfo_children():
            widget.destroy()
        
        # Botón de cancelar
        tk.Button(self.control_frame, text="❌ Cancelar", 
                 command=self.cancel_scraping, bg=STYLE["danger"], fg="white",
                 font=("Segoe UI", 9, "bold"), relief="flat", padx=15).pack(side="left", padx=5)
        
        self.control_frame.pack(pady=5)
        
        # Iniciar thread de scraping optimizado
        self.scraper_thread = FastScraperThread(
            login_data,
            self.on_progress_update,
            self.on_captcha_required,
            self.on_scraping_finished
        )
        self.scraper_thread.start()
    
    def show_progress_frame(self):
        """Muestra el frame de progreso"""
        self.progress_frame.pack(fill="x", before=self.body)
        self.progress_frame.pack_propagate(False)
    
    def hide_progress_frame(self):
        """Oculta el frame de progreso"""
        self.progress_frame.pack_forget()
        # Limpiar botones de control
        for widget in self.control_frame.winfo_children():
            widget.destroy()
        self.control_frame.pack_forget()
    
    def cancel_scraping(self):
        """Cancela el proceso de scraping"""
        if self.scraper_thread and self.scraper_thread.is_alive():
            if messagebox.askyesno("Cancelar", "¿Estás seguro de cancelar el scraping?"):
                self.scraper_thread.stop()
                self.hide_progress_frame()
                messagebox.showinfo("Cancelado", "Scraping cancelado.")
    
    def on_progress_update(self, message, msg_type="info"):
        """Callback para actualizar progreso"""
        self.progress_queue.put(("progress", message, msg_type))
    
    def on_captcha_required(self, captcha_image_data, driver=None):
        """Callback cuando se requiere CAPTCHA"""
        if captcha_image_data == "manual":
            # Modo manual
            self.captcha_window = CaptchaViewerWindow(self.root, driver)
            self.root.wait_window(self.captcha_window)
            return self.captcha_window.solved
        elif captcha_image_data == "check":
            # Verificar si ya se resolvió manualmente
            if self.captcha_window:
                return self.captcha_window.solved
            return False
    
    def on_scraping_finished(self, success, result):
        """Callback cuando termina el scraping"""
        self.progress_queue.put(("finished", success, result))
    
    def _start_queue_processor(self):
        """Procesa mensajes de las colas"""
        try:
            # Procesar cola de progreso
            while True:
                try:
                    msg_type, *args = self.progress_queue.get_nowait()
                    if msg_type == "progress":
                        message, msg_type = args
                        colors = {
                            "info": STYLE["info"],
                            "success": STYLE["success"],
                            "warning": STYLE["warning"],
                            "error": STYLE["danger"]
                        }
                        self.progress_label.config(text=message, fg=colors.get(msg_type, "white"))
                    elif msg_type == "finished":
                        success, result = args
                        self.hide_progress_frame()
                        if success:
                            messagebox.showinfo("Éxito", f"Scraping completado.\nArchivo: {result}")
                            self.load_csv_file(result)
                        else:
                            messagebox.showerror("Error", f"Scraping falló: {result}")
                except queue.Empty:
                    break
        except:
            pass
            
        self.root.after(100, self._start_queue_processor)
    
    # --- MÉTODOS PARA CARGAR Y PROCESAR CSV CORRECTAMENTE ---
    
    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path: 
            return
        self.load_csv_file(path)
    
    def load_csv_file(self, path):
        """Carga un archivo CSV y lo procesa correctamente"""
        try:
            # RESET
            self.inscritas.clear()
            self.materias_inscritas.clear()
            self.f_materia.set("")
            self.f_docente.set("")
            self.f_dia.set("")
            self.count_var.set("Materias Registradas: 0")
            
            # Leer CSV con manejo de errores
            try:
                self.df = pd.read_csv(path, encoding='utf-8')
            except UnicodeDecodeError:
                self.df = pd.read_csv(path, encoding='latin-1')
            
            # Verificar que tiene las columnas necesarias
            required_columns = ['Materia', 'Docente', 'Paralelo', 'Modalidad']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                messagebox.showerror("Error", f"El archivo CSV no tiene las columnas necesarias.\nFaltan: {', '.join(missing_columns)}")
                return
            
            # Limpiar datos
            self.df = self.df.fillna('')
            
            # Convertir columnas de texto
            self.df['Materia'] = self.df['Materia'].astype(str).str.strip()
            self.df['Docente'] = self.df['Docente'].astype(str).str.strip()
            
            # Crear columna uid
            self.df['uid'] = range(len(self.df))
            
            # Preparar vista
            self.df_view = self.df.copy()
            
            # Generar iniciales únicas para cada materia
            self._generate_unique_initials()
            
            # Actualizar comboboxes
            self.update_comboboxes()
            self.cb_sug_doc['values'] = ["Cualquiera"] + sorted(self.df['Docente'].astype(str).unique())
            
            # Asignar colores
            for m in self.df['Materia'].unique(): 
                self._get_color(m)
            
            # Actualizar interfaz
            self.refresh_grid()
            self.scan_global_warnings()
            self.generate_suggestions()
            
            # Mostrar estadísticas
            num_materias = len(self.df['Materia'].unique())
            num_paralelos = len(self.df)
            messagebox.showinfo("Carga Completa", 
                              f"Horario cargado exitosamente.\n"
                              f"• Materias: {num_materias}\n"
                              f"• Paralelos: {num_paralelos}")
                              
        except Exception as e: 
            messagebox.showerror("Error", f"No se pudo cargar el archivo CSV:\n{str(e)}")
            print(f"Error detallado: {e}")
    
    def _generate_unique_initials(self):
        """Genera iniciales únicas para cada materia: 2 primeros caracteres de cada palabra"""
        self.initials_map = {}
        used_initials = set()
        
        for materia in self.df['Materia'].unique():
            # Separar en palabras y tomar los dos primeros caracteres de cada una
            palabras = materia.split()
            # Tomar los dos primeros caracteres de cada palabra, si tiene menos de 2, tomar la palabra completa
            candidato = ''.join([palabra[:2] for palabra in palabras])
            # Si el candidato está vacío (no hay palabras), usar los primeros dos caracteres de la materia
            if not candidato:
                candidato = materia[:2]
            
            # Asegurar que sea único
            original_candidate = candidato
            counter = 1
            while candidato in used_initials:
                candidato = f"{original_candidate}{counter}"
                counter += 1
            
            self.initials_map[materia] = candidato.upper()
            used_initials.add(candidato)
    
    def get_initials(self, materia):
        """Obtiene las iniciales únicas para una materia"""
        return self.initials_map.get(materia, materia[:4].upper())
    
    def update_comboboxes(self):
        """Actualiza los comboboxes de filtros"""
        if self.df.empty:
            self.cb_materia['values'] = []
            self.cb_docente['values'] = []
            self.cb_dia['values'] = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
            return
            
        df_base = self.df.copy()
        m, d = self.f_materia.get(), self.f_docente.get()
        
        if d: 
            mats = sorted(df_base[df_base['Docente'] == d]['Materia'].astype(str).unique())
        else: 
            mats = sorted(df_base['Materia'].astype(str).unique())
        
        if m: 
            docs = sorted(df_base[df_base['Materia'] == m]['Docente'].astype(str).unique())
        else: 
            docs = sorted(df_base['Docente'].astype(str).unique())
        
        self.cb_materia['values'] = mats
        self.cb_docente['values'] = docs
        self.cb_dia['values'] = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    
    def on_filter_change(self, event):
        """Aplica los filtros seleccionados"""
        if self.df.empty:
            return
            
        df = self.df.copy()
        m, d, day = self.f_materia.get(), self.f_docente.get(), self.f_dia.get()
        
        if m: 
            df = df[df['Materia'] == m]
        if d: 
            df = df[df['Docente'] == d]
        if day: 
            mask = (df['Dia_T1'] == day) | (df['Dia_T2'] == day) | (df['Dia_P'] == day)
            df = df[mask]
        
        self.df_view = df
        self.update_comboboxes()
        self.refresh_grid()
    
    # --- MÉTODOS PARA GENERAR HORARIOS ---
    
    def _get_color(self, materia):
        if materia not in self.color_map:
            random.seed(materia)
            h = random.random()
            import colorsys
            r, g, b = colorsys.hsv_to_rgb(h, 0.6, 0.85) 
            self.color_map[materia] = f'#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}'
        return self.color_map[materia]
    
    def format_name(self, full):
        try:
            p = str(full).split()
            if len(p) >= 3: 
                return f"{p[2].title()} {p[0].title()}"
            elif len(p) == 2: 
                return f"{p[1].title()} {p[0].title()}"
            return full.title()
        except: 
            return str(full)
    
    def get_slots(self, row):
        mapa = {"Lunes": 0, "Martes": 1, "Miércoles": 2, "Jueves": 3, "Viernes": 4}
        slots = []
        cols = [('Dia_T1', 'HoraI_T1', 'HoraF_T1'), ('Dia_T2', 'HoraI_T2', 'HoraF_T2'), ('Dia_P', 'HoraI_P', 'HoraF_P')]
        
        for cd, hi, hf in cols:
            dia = row.get(cd, '')
            if pd.notna(dia) and dia in mapa:
                try:
                    # Procesar hora de inicio
                    hora_inicio_str = str(row.get(hi, '')).strip()
                    hora_fin_str = str(row.get(hf, '')).strip()
                    
                    if hora_inicio_str and hora_fin_str:
                        # Extraer solo la parte de la hora (ignorar minutos si es necesario)
                        s = int(hora_inicio_str.split(':')[0])
                        e = int(hora_fin_str.split(':')[0])
                        
                        for h in range(s, e): 
                            slots.append({'dia': mapa[dia], 'hora': h, 'start': s})
                except (ValueError, AttributeError) as e:
                    # Silenciar errores de conversión
                    pass
        return slots
    
    def refresh_grid(self):
        # Limpiar grilla
        for f in self.cell_frames.values():
            for w in f.winfo_children(): 
                w.destroy()
            f.config(bg=STYLE["bg_cell_empty"], cursor="arrow")
        
        if self.df_view.empty:
            return
        
        # Disponibles
        avail = {}
        for _, row in self.df_view.iterrows():
            if row['Materia'] in self.materias_inscritas and row['uid'] not in self.inscritas: 
                continue
            for s in self.get_slots(row):
                k = (s['dia'], s['hora'])
                if k not in avail: 
                    avail[k] = set()
                avail[k].add((row['Materia'], row['Paralelo'], row['uid']))
        
        for (d, h), opts in avail.items():
            if (d, h) not in self.cell_frames: 
                continue
            
            # Verificar si está ocupado
            occupied = False
            for uid in self.inscritas:
                slots = self.get_slots(self.df.loc[uid])
                if any(x['dia'] == d and x['hora'] == h for x in slots):
                    occupied = True
                    break
            
            if occupied: 
                continue
            
            frame = self.cell_frames[(d, h)]
            frame.config(cursor="hand2")
            
            # Contar materias únicas y paralelos
            materias_unicas = set([m for m, p, u in opts])
            num_materias = len(materias_unicas)
            num_paralelos = len(opts)
            
            # Mostrar badges de iniciales en la esquina superior derecha - MÁS MINIMALISTA
            badge_frame = tk.Frame(frame, bg=STYLE["bg_cell_empty"])
            badge_frame.place(relx=1.0, rely=0.0, anchor="ne", x=-1, y=1)
            
            # Mostrar hasta 3 materias como badges apilados
            materias_lista = sorted(materias_unicas)
            for i, materia in enumerate(materias_lista[:3]):
                color = self._get_color(materia)
                initials = self.get_initials(materia)
                
                # Contar cuántos paralelos tiene esta materia en esta celda
                paralelos_materia = len([p for m, p, u in opts if m == materia])
                
                # Crear badge minimalista
                badge = tk.Frame(badge_frame, bg=color)
                badge.pack(pady=0.5)
                
                # Mostrar iniciales
                tk.Label(badge, text=initials, bg=color, fg="black", 
                        font=("Arial", 5), padx=1, pady=0).pack()
            
            # Si hay más de 3 materias, mostrar badge con "+X"
            if num_materias > 3:
                badge = tk.Frame(badge_frame, bg="#444")
                badge.pack(pady=0.5)
                tk.Label(badge, text=f"+{num_materias-3}", bg="#444", fg="white", 
                        font=("Arial", 5), padx=1, pady=0).pack()
            
            # Mostrar el total de opciones en el centro
            center_label = tk.Label(frame, text=f"{num_paralelos} opc", 
                                  bg=STYLE["bg_cell_empty"], fg=STYLE["avail"],
                                  font=("Segoe UI", 8, "bold"))
            center_label.place(relx=0.5, rely=0.5, anchor="center")
            
            # Hacer clickable
            center_label.bind("<Button-1>", lambda e, dd=d, hh=h: self.on_cell_click(dd, hh))
            badge_frame.bind("<Button-1>", lambda e, dd=d, hh=h: self.on_cell_click(dd, hh))
            frame.bind("<Button-1>", lambda e, dd=d, hh=h: self.on_cell_click(dd, hh))
        
        # Mostrar materias inscritas
        for uid in self.inscritas:
            row = self.df.loc[uid]
            col = self._get_color(row['Materia'])
            doc = self.format_name(row['Docente'])
            
            for s in self.get_slots(row):
                if (s['dia'], s['hora']) in self.cell_frames:
                    frm = self.cell_frames[(s['dia'], s['hora'])]
                    frm.config(bg=col, cursor="hand2")
                    
                    # Badge con iniciales en esquina superior derecha - MINIMALISTA
                    initials = self.get_initials(row['Materia'])
                    badge_bg = col
                    badge_fg = "#000" if sum(int(col[i:i+2], 16) for i in (1, 3, 5)) > 382 else "#fff"
                    
                    badge = tk.Frame(frm, bg=badge_bg)
                    badge.place(relx=1.0, rely=0.0, anchor="ne", x=-1, y=1)
                    tk.Label(badge, text=initials, bg=badge_bg, fg=badge_fg,
                            font=("Arial", 6), padx=1, pady=0).pack()
                    
                    if s['hora'] == s['start']:
                        c = tk.Frame(frm, bg=col)
                        c.place(relx=0.5, rely=0.5, anchor="center")
                        
                        tk.Label(c, text=row['Materia'][:15] + ("..." if len(row['Materia']) > 15 else ""), 
                                bg=col, fg="#000", font=("Segoe UI", 8, "bold"), 
                                wraplength=130).pack()
                        tk.Label(c, text=doc[:15] + ("..." if len(doc) > 15 else ""), 
                                bg=col, fg="#222", font=("Segoe UI", 7)).pack()
                        
                        aula = row.get('Aula_T1', '') or row.get('Aula_P', '') or '?'
                        tk.Label(c, text=f"P{row['Paralelo']} | {aula[:10]}", 
                                bg=col, fg="#333", font=("Arial", 6)).pack()
                        
                        for w in c.winfo_children(): 
                            w.bind("<Button-1>", lambda e, dd=s['dia'], hh=s['hora']: self.on_cell_click(dd, hh))
                        c.bind("<Button-1>", lambda e, dd=s['dia'], hh=s['hora']: self.on_cell_click(dd, hh))
    
    def on_cell_click(self, d, h):
        """Maneja clic en celda de horario"""
        # Limpiar panel de detalles
        for w in self.frame_list.winfo_children(): 
            w.destroy()
        
        dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
        tk.Label(self.frame_list, text=f"{dias[d]} - {h}:00", bg=STYLE["bg_panel"], 
                fg="white", font=STYLE["font_h"]).pack(pady=10, anchor="w")
        
        # Buscar materia inscrita en esta celda
        inscrita = None
        for uid in self.inscritas:
            row = self.df.loc[uid]
            if any(s['dia'] == d and s['hora'] == h for s in self.get_slots(row)):
                inscrita = row
                break
        
        if inscrita is not None:
            tk.Label(self.frame_list, text="ESTADO: REGISTRADA", bg=STYLE["bg_panel"], 
                    fg=STYLE["success"], font=("Segoe UI", 9, "bold")).pack(anchor="w")
            self._render_card(inscrita, "inscrito", self.frame_list)
            tk.Frame(self.frame_list, height=1, bg="#444").pack(fill="x", pady=15)
            tk.Label(self.frame_list, text="Otras opciones:", bg=STYLE["bg_panel"], 
                    fg="#888").pack(anchor="w")
        
        # Buscar otras opciones en esta celda
        seen = set()
        found = False
        
        for _, row in self.df.iterrows():
            slots = self.get_slots(row)
            if any(s['dia'] == d and s['hora'] == h for s in slots):
                key = (row['Materia'], row['Paralelo'])
                if key in seen: 
                    continue
                seen.add(key)
                
                # Si es la fila inscrita, saltar
                if inscrita is not None and row['uid'] == inscrita['uid']: 
                    continue
                
                state = "available"
                conflict, reason = self.check_conflict_detailed(row, ignore_subject=row['Materia'])
                
                if row['Materia'] in self.materias_inscritas and inscrita is None: 
                    state = "switch"
                    if conflict: 
                        state = "conflict"
                    curr_p = "?"
                    for u in self.inscritas:
                        if self.df.loc[u]['Materia'] == row['Materia']: 
                            curr_p = self.df.loc[u]['Paralelo']
                    self._render_card(row, state, self.frame_list, reason, extra_info=f"Tienes P{curr_p}")
                
                elif inscrita is not None:
                    # Si es la misma materia (switch)
                    if row['Materia'] == inscrita['Materia']:
                        state = "switch"
                        self._render_card(row, state, self.frame_list, reason, extra_info=f"Tienes P{inscrita['Paralelo']}")
                    else:
                        state = "blocked"
                        self._render_card(row, state, self.frame_list)
                
                elif conflict: 
                    state = "conflict"
                    self._render_card(row, state, self.frame_list, reason)
                else: 
                    self._render_card(row, state, self.frame_list, reason)
                found = True
        
        if not found and inscrita is None: 
            tk.Label(self.frame_list, text="No hay opciones.", bg=STYLE["bg_panel"], fg="#666").pack()
        
        if not self.pane_detalles.expanded: 
            self.pane_detalles.toggle()
    
    def _render_card(self, row, state, parent_frame, conflict_reason=None, extra_info=""):
        bg = "#2a2a2a" if state != "inscrito" else "#203020"
        card = tk.Frame(parent_frame, bg=bg, pady=8, padx=8)
        card.pack(fill="x", pady=5)
        
        # Barra de color
        tk.Frame(card, bg=self._get_color(row['Materia']), width=4).pack(side="left", fill="y")
        
        # Información
        info = tk.Frame(card, bg=bg, padx=10)
        info.pack(side="left", fill="both", expand=True)
        
        # Mostrar iniciales junto al nombre
        initials = self.get_initials(row['Materia'])
        name_frame = tk.Frame(info, bg=bg)
        name_frame.pack(fill="x", pady=(0, 2))
        
        # Badge con iniciales - MINIMALISTA
        badge_bg = self._get_color(row['Materia'])
        badge_fg = "#000" if sum(int(badge_bg[i:i+2], 16) for i in (1, 3, 5)) > 382 else "#fff"
        badge = tk.Frame(name_frame, bg=badge_bg)
        badge.pack(side="left", padx=(0, 5))
        tk.Label(badge, text=initials, bg=badge_bg, fg=badge_fg,
                font=("Arial", 7), padx=1, pady=0).pack()
        
        tk.Label(name_frame, text=row['Materia'][:25] + ("..." if len(row['Materia']) > 25 else ""), 
                font=("Segoe UI", 9, "bold"), bg=bg, fg="white", anchor="w").pack(side="left", fill="x", expand=True)
        
        tk.Label(info, text=self.format_name(row['Docente']), font=STYLE["font_txt"], 
                bg=bg, fg="#ccc", anchor="w").pack(fill="x")
        
        # Mostrar aula si existe
        aula = row.get('Aula_T1', '') or row.get('Aula_P', '') or 'N/A'
        tk.Label(info, text=f"Aula: {aula}", font=STYLE["font_small"], 
                bg=bg, fg="#ddd", anchor="w").pack(fill="x")
        
        tk.Label(info, text=f"P{row['Paralelo']} | {row.get('Modalidad', 'N/A')}", 
                font=STYLE["font_small"], bg=bg, fg="#888", anchor="w").pack(fill="x")
        
        # Botones según estado
        if state == "inscrito":
            tk.Button(card, text="RETIRAR", bg=STYLE["danger"], fg="white", 
                     relief="flat", font=("Segoe UI", 7, "bold"), 
                     command=lambda: self.toggle_enroll(row)).pack(side="right", padx=5)
        elif state == "blocked":
            tk.Button(card, text="⛔ OCUPADO", bg="#444", fg="#888", 
                     relief="flat", font=("Segoe UI", 7), state="disabled").pack(side="right", padx=5)
        elif state == "switch":
            btn_box = tk.Frame(card, bg=bg)
            btn_box.pack(side="right", padx=5)
            tk.Button(btn_box, text="🔄 CAMBIAR", bg=STYLE["switch"], fg="black", 
                     relief="flat", font=("Segoe UI", 7, "bold"),
                     command=lambda: self.switch_enroll(row)).pack(anchor="e")
            if extra_info: 
                tk.Label(btn_box, text=extra_info, bg=bg, fg="#aaa", 
                        font=("Arial", 6)).pack(anchor="e")
        elif state == "conflict":
            btn_box = tk.Frame(card, bg=bg)
            btn_box.pack(side="right", padx=5)
            tk.Button(btn_box, text="⚠ CONFLICTO", bg=STYLE["warning"], fg="black", 
                     relief="flat", font=("Segoe UI", 7, "bold"),
                     command=lambda: messagebox.showwarning("Detalle del Conflicto", conflict_reason)).pack(anchor="e")
        else:
            tk.Button(card, text="INSCRIBIR", bg=STYLE["avail"], fg="black", 
                     relief="flat", font=("Segoe UI", 7, "bold"), 
                     command=lambda: self.toggle_enroll(row)).pack(side="right", padx=5)
    
    def check_conflict_detailed(self, candidate_row, ignore_subject=None):
        cand_slots = {(s['dia'], s['hora']) for s in self.get_slots(candidate_row)}
        for uid in self.inscritas:
            ex_row = self.df.loc[uid]
            if ignore_subject and ex_row['Materia'] == ignore_subject: 
                continue
            
            # 1. Clases
            ex_slots = {(s['dia'], s['hora']) for s in self.get_slots(ex_row)}
            if cand_slots & ex_slots: 
                d_idx = list(cand_slots & ex_slots)[0][0]
                dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
                return True, f"Cruce de Clases:\n{candidate_row['Materia']} choca con {ex_row['Materia']} el {dias[d_idx]}."
            
            # 2. Exámenes (si existen)
            for c in ['Examen1', 'Examen2', 'Examen3']:
                p1 = self.parse_exam_datetime(candidate_row.get(c))
                p2 = self.parse_exam_datetime(ex_row.get(c))
                if p1 and p2:
                    if p1[0] == p2[0] and max(p1[1], p2[1]) < min(p1[2], p2[2]):
                         return True, f"Cruce de Examen ({p1[0]}):\n{candidate_row['Materia']} vs {ex_row['Materia']}"
        return False, None
    
    def parse_exam_datetime(self, exam_str):
        try:
            parts = str(exam_str).split(' - ')
            if len(parts) < 2: 
                return None
            date_part = parts[0].strip()
            time_part = parts[1].strip() 
            start_str, end_str = time_part.split(' a ')
            def to_min(t):
                h, m = map(int, t.split(':'))
                return h * 60 + m
            return date_part, to_min(start_str), to_min(end_str)
        except: 
            return None
    
    def switch_enroll(self, new_row):
        old_uid = None
        for uid in self.inscritas:
            if self.df.loc[uid]['Materia'] == new_row['Materia']: 
                old_uid = uid
                break
        if old_uid is not None: 
            self.inscritas.remove(old_uid)
        self.inscritas.add(new_row['uid'])
        self.reset_filters()
        self.refresh_grid()
        self.generate_suggestions()
        try:
            s = self.get_slots(new_row)[0]
            self.on_cell_click(s['dia'], s['hora'])
        except: 
            pass
    
    def toggle_enroll(self, row):
        uid = row['uid']
        if uid in self.inscritas:
            self.inscritas.remove(uid)
            self.materias_inscritas.remove(row['Materia'])
        else:
            if row['Materia'] in self.materias_inscritas:
                messagebox.showwarning("Duplicado", f"Ya estás inscrito en {row['Materia']}.")
                return
            conflict, reason = self.check_conflict_detailed(row)
            if conflict: 
                messagebox.showerror("Conflicto", reason)
                return
            self.inscritas.add(uid)
            self.materias_inscritas.add(row['Materia'])
            self.reset_filters()
        
        self.count_var.set(f"Materias Registradas: {len(self.materias_inscritas)}")
        self.refresh_grid()
        self.generate_suggestions()
        try:
            s = self.get_slots(row)[0]
            self.on_cell_click(s['dia'], s['hora'])
        except: 
            pass
    
    def generate_suggestions(self):
        """Genera sugerencias compatibles"""
        for w in self.frame_sug.winfo_children(): 
            w.destroy()
        if self.df.empty: 
            return
            
        tk.Label(self.frame_sug, text="Sugerencias sin conflictos de horario:", 
                bg=STYLE["bg_panel"], fg=STYLE["avail"], font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0,10))
        
        candidates = []
        seen_sug = set()
        
        pref_turno = self.sug_turno_var.get()
        pref_doc = self.sug_docente_var.get()
        
        shuffled_df = self.df.sample(frac=1).reset_index(drop=True)
        for _, row in shuffled_df.iterrows():
            if len(candidates) >= 5: 
                break 
            if row['Materia'] in self.materias_inscritas: 
                continue
            if (row['Materia'], row['Paralelo']) in seen_sug: 
                continue
            
            if pref_doc and pref_doc != "Cualquiera" and pref_doc.lower() not in str(row['Docente']).lower(): 
                continue
            
            slots = self.get_slots(row)
            if not slots: 
                continue
            start_h = slots[0]['start']
            
            match_turno = True
            if pref_turno == "Mañana (<12)" and start_h >= 12: 
                match_turno = False
            elif pref_turno == "Tarde (12-17)" and (start_h < 12 or start_h >= 17): 
                match_turno = False
            elif pref_turno == "Noche (>17)" and start_h < 17: 
                match_turno = False
            if not match_turno: 
                continue
            
            conflict, reason = self.check_conflict_detailed(row)
            if conflict:
                candidates.append((row, "conflict", reason))
                seen_sug.add((row['Materia'], row['Paralelo']))
            else:
                candidates.append((row, "available", None))
                seen_sug.add((row['Materia'], row['Paralelo']))
        
        if not candidates: 
            tk.Label(self.frame_sug, text="Sin sugerencias disponibles.", bg=STYLE["bg_panel"], fg="#666").pack(pady=20)
            return
            
        available_candidates = [c for c in candidates if c[1] == "available"]
        conflict_candidates = [c for c in candidates if c[1] == "conflict"]
        
        if available_candidates:
            tk.Label(self.frame_sug, text="✅ Compatibles:", 
                    bg=STYLE["bg_panel"], fg=STYLE["success"], font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5,0))
            for row, state, reason in available_candidates: 
                self._render_card(row, state, self.frame_sug, reason)
        
        if conflict_candidates:
            tk.Label(self.frame_sug, text="⚠ Con conflicto:", 
                    bg=STYLE["bg_panel"], fg=STYLE["warning"], font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(10,0))
            for row, state, reason in conflict_candidates: 
                self._render_card(row, state, self.frame_sug, reason)
    
    def scan_global_warnings(self):
        """Escanear conflictos globales sin duplicados"""
        self.txt_global.config(state="normal")
        self.txt_global.delete(1.0, "end")
        
        all_exams = []
        exam_details = {}  # Para evitar duplicados
        
        for idx, row in self.df.iterrows():
            mat = row['Materia']
            par = row['Paralelo']
            doc = self.format_name(row['Docente'])
            
            for c in ['Examen1', 'Examen2', 'Examen3']:
                exam_value = row.get(c)
                if pd.notna(exam_value) and str(exam_value).strip():
                    p = self.parse_exam_datetime(exam_value)
                    if p:
                        # Crear clave única para evitar duplicados
                        exam_key = (mat, par, doc, c, p[0], p[1], p[2])
                        if exam_key not in exam_details:
                            exam_details[exam_key] = {
                                'mat': mat, 'par': par, 'doc': doc, 'exam_type': c,
                                'date': p[0], 'start': p[1], 'end': p[2],
                                'time_str': str(exam_value).split(' - ')[1] if ' - ' in str(exam_value) else "?",
                                'uid': idx
                            }
                            all_exams.append(exam_details[exam_key])
        
        found = False
        reported_pairs = set()
        
        for i in range(len(all_exams)):
            for j in range(i + 1, len(all_exams)):
                e1 = all_exams[i]
                e2 = all_exams[j]
                
                if e1['mat'] == e2['mat']: 
                    continue 
                
                if e1['date'] == e2['date']:
                    if max(e1['start'], e2['start']) < min(e1['end'], e2['end']):
                        # Usar clave ordenada para evitar duplicados
                        pair_key = tuple(sorted([(e1['mat'], e1['par'], e1['exam_type']), 
                                                (e2['mat'], e2['par'], e2['exam_type'])]))
                        
                        if pair_key not in reported_pairs:
                            reported_pairs.add(pair_key)
                            found = True
                            
                            # Formatear información sin duplicados
                            time_display = f"{e1['time_str']}" if e1['time_str'] == e2['time_str'] else f"{e1['time_str']} / {e2['time_str']}"
                            
                            msg = (
                                f"⚠ CRUCE DE EXÁMENES: {e1['date']} @ {time_display}\n"
                                f"   • {e1['mat']} (P{e1['par']}) - {e1['doc']} [{e1['exam_type']}]\n"
                                f"   • {e2['mat']} (P{e2['par']}) - {e2['doc']} [{e2['exam_type']}]\n"
                                f"{'-'*40}\n"
                            )
                            self.txt_global.insert("end", msg, "warn")
        
        if not found: 
            self.txt_global.insert("end", "✔ Sin conflictos globales detectados.", "ok")
        
        self.txt_global.tag_config("warn", foreground=STYLE["warning"])
        self.txt_global.tag_config("ok", foreground=STYLE["success"])
        self.txt_global.config(state="disabled")
    
    def on_key_release(self, event, variable, cbox):
        typed = cbox.get()
        if typed == '': 
            self.update_comboboxes()
            return
        if cbox == getattr(self, 'cb_materia', None): 
            full = sorted(self.df['Materia'].astype(str).unique())
        elif cbox == getattr(self, 'cb_docente', None): 
            full = sorted(self.df['Docente'].astype(str).unique())
        else: 
            full = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
        filtered = [i for i in full if typed.lower() in i.lower()]
        cbox['values'] = filtered
        if filtered: 
            cbox.event_generate('<Down>')
    
    def on_sug_doc_key(self, event):
        typed = self.cb_sug_doc.get()
        full = ["Cualquiera"] + sorted(self.df['Docente'].astype(str).unique()) if not self.df.empty else []
        filtered = [i for i in full if typed.lower() in i.lower()]
        self.cb_sug_doc['values'] = filtered
        if filtered: 
            self.cb_sug_doc.event_generate('<Down>')
    
    def _build_grid_structure(self):
        dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
        horas = range(7, 21)
        tk.Label(self.grid_container, text="HORA", bg=STYLE["bg_app"], fg="#666", 
                font=("Segoe UI", 9, "bold")).grid(row=0, column=0, pady=5)
        for i, d in enumerate(dias):
            tk.Label(self.grid_container, text=d.upper(), bg=STYLE["bg_app"], 
                    fg=STYLE["fg_prim"], font=("Segoe UI", 10, "bold")).grid(row=0, column=i+1, sticky="ew")
            self.grid_container.columnconfigure(i+1, weight=1)
        for r, h in enumerate(horas):
            tk.Label(self.grid_container, text=f"{h:02d}:00", bg=STYLE["bg_app"], 
                    fg="#555", font=("Consolas", 9)).grid(row=r+1, column=0)
            for c, dia in enumerate(dias):
                frame = tk.Frame(self.grid_container, bg=STYLE["bg_cell_empty"], bd=1, relief="solid")
                frame.grid(row=r+1, column=c+1, sticky="nsew", padx=1, pady=1)
                frame.bind("<Button-1>", lambda e, d=c, hr=h: self.on_cell_click(d, hr))
                self.cell_frames[(c, h)] = frame
            self.grid_container.rowconfigure(r+1, weight=1)
    
    def open_exam_window(self):
        if not self.inscritas: 
            messagebox.showinfo("Vacío", "No tienes materias inscritas.")
            return
            
        top = tk.Toplevel(self.root)
        top.title("Horario de Exámenes")
        top.geometry("1000x600")
        top.configure(bg=STYLE["bg_app"])
        
        main_canvas = tk.Canvas(top, bg=STYLE["bg_app"], highlightthickness=0)
        scroll_y = ttk.Scrollbar(top, orient="vertical", command=main_canvas.yview)
        scroll_frame = tk.Frame(main_canvas, bg=STYLE["bg_app"])
        scroll_frame.bind("<Configure>", lambda e: main_canvas.configure(
            scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=960)
        main_canvas.configure(yscrollcommand=scroll_y.set)
        main_canvas.pack(side="left", fill="both", expand=True, padx=10)
        scroll_y.pack(side="right", fill="y")
        
        tk.Label(scroll_frame, text="MATERIAS INSCRITAS", bg=STYLE["bg_app"], 
                fg=STYLE["success"], font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10,5))
        
        container_ok = tk.Frame(scroll_frame, bg=STYLE["bg_app"])
        container_ok.pack(fill="x", pady=5)
        
        headers = ["Materia", "1er Parcial", "2do Parcial", "3er Parcial"]
        for i, h in enumerate(headers): 
            tk.Label(container_ok, text=h, bg="#333", fg="white", 
                    font=("Segoe UI", 9, "bold"), pady=5).grid(row=0, column=i, sticky="ew", padx=2)
            container_ok.columnconfigure(i, weight=1)
        
        rows = [self.df.loc[uid] for uid in self.inscritas]
        for idx, row in enumerate(rows):
            bg = STYLE["bg_cell_empty"] if idx % 2 == 0 else "#252526"
            color = self._get_color(row['Materia'])
            
            m_frame = tk.Frame(container_ok, bg=bg)
            m_frame.grid(row=idx+1, column=0, sticky="ew", pady=1, padx=2)
            tk.Frame(m_frame, bg=color, width=5).pack(side="left", fill="y")
            
            # Mostrar iniciales junto al nombre
            name_frame = tk.Frame(m_frame, bg=bg)
            name_frame.pack(side="left", fill="x", expand=True)
            
            # Badge con iniciales - MINIMALISTA
            badge_bg = color
            badge_fg = "#000" if sum(int(badge_bg[i:i+2], 16) for i in (1, 3, 5)) > 382 else "#fff"
            badge = tk.Frame(name_frame, bg=badge_bg)
            badge.pack(side="left", padx=(5, 5))
            tk.Label(badge, text=self.get_initials(row['Materia']), bg=badge_bg, fg=badge_fg,
                    font=("Arial", 7), padx=1, pady=0).pack()
            
            tk.Label(name_frame, text=row['Materia'], bg=bg, fg="white", 
                    font=("Segoe UI", 9, "bold"), padx=5).pack(side="left")
            
            exam_cols = [('Examen1', 'Aula1'), ('Examen2', 'Aula2'), ('Examen3', 'Aula3')]
            for c_idx, (col_date, col_room) in enumerate(exam_cols):
                d_val = str(row.get(col_date, 'N/A'))
                r_val = str(row.get(col_room, '')) if col_room in row else ""
                txt_display = d_val + (f"\n({r_val})" if r_val and r_val.lower() != "nan" else "")
                tk.Label(container_ok, text=txt_display, bg=bg, fg="#ccc", 
                        font=("Consolas", 8), justify="center").grid(row=idx+1, column=c_idx+1, sticky="ew", pady=1, padx=2)
    
    def save_screenshot(self):
        if not HAS_PILLOW: 
            messagebox.showwarning("Falta Pillow", "Instala: pip install pillow")
            return
        try:
            x = self.grid_container.winfo_rootx()
            y = self.grid_container.winfo_rooty()
            w = self.grid_container.winfo_width()
            h = self.grid_container.winfo_height()
            img = ImageGrab.grab(bbox=(x, y, x+w, y+h))
            path = filedialog.asksaveasfilename(defaultextension=".png", 
                                               initialfile="Horario_Academico.png", 
                                               filetypes=[("PNG", "*.png")])
            if path: 
                img.save(path)
                messagebox.showinfo("Guardado", f"Imagen guardada en:\n{path}")
        except Exception as e: 
            messagebox.showerror("Error", str(e))
    
    def reset_filters(self):
        self.f_materia.set("")
        self.f_docente.set("")
        self.f_dia.set("")
        self.update_comboboxes()
        self.on_filter_change(None)
    
    def clear_schedule(self):
        if messagebox.askyesno("Limpiar", "¿Borrar todo tu horario?"):
            self.inscritas.clear()
            self.materias_inscritas.clear()
            self.count_var.set("Materias Registradas: 0")
            self.reset_filters()
            self.generate_suggestions()
            messagebox.showinfo("Limpio", "Horario borrado")
    
    def on_close(self):
        """Manejar cierre de aplicación"""
        if self.scraper_thread and self.scraper_thread.is_alive():
            self.scraper_thread.stop()
            self.scraper_thread.join(timeout=2)
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AcademicScheduler(root)
    
    # Configurar cierre limpio
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    
    # Aplicar tema oscuro
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configurar colores para widgets ttk
    style.configure('TCombobox', fieldbackground='#2a2a2a', background='#2a2a2a', 
                   foreground='white', selectbackground=STYLE["accent"])
    
    root.mainloop()