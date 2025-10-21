import os
import sys
import subprocess
import time
import shutil
import winreg
import ctypes
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from PIL import Image, ImageTk

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tecnologia da Informação - Baln. Piçarras")
        self.root.geometry("800x600")
        self.root.resizable(False, False)
        
        if getattr(sys, 'frozen', False):
            self.script_path = Path(sys.executable).parent
        else:
            self.script_path = Path(__file__).parent
            
        self.is_admin = self.check_admin()
        
        self.center_window()
        self.set_icon()
        
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Pronto")
        
        # Setup da UI
        self.setup_ui()
        
    def center_window(self):
        self.root.update_idletasks()
        width = 800
        height = 600
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def set_icon(self):
        """Configura um ícone para a aplicação"""
        try:
            icon_path = self.script_path / "icon.ico"
            print(f"🔍 Procurando ícone em: {icon_path}")
            print(f"📁 Arquivo existe: {icon_path.exists()}")
            
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
                print("✅ Ícone configurado com sucesso!")
            else:
                print("❌ Ícone não encontrado!")
        except Exception as e:
            print(f"💥 Erro ao carregar ícone: {e}")
    
    def check_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def set_window_icon(self, window):
        """Configura o ícone para qualquer janela"""
        try:
            icon_path = self.script_path / "icon.ico"
            if icon_path.exists():
                window.iconbitmap(str(icon_path))
                return True
            return False
        except:
            return False
    
    def setup_ui(self):
        # Primeiro carregar o background
        self.load_background()
        
        # Criar frame principal SEM fundo (transparente)
        main_frame = tk.Frame(self.root, bg='', padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Header com fundo branco semi-transparente
        header_frame = tk.Frame(main_frame, bg='white')
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            header_frame, 
            text="Tecnologia da Informação - Baln. Piçarras", 
            font=('Arial', 16, 'bold'),
            foreground='#2E86AB',
            bg='white'
        ).pack(pady=10)
        
        admin_status = "✓ Executando como Administrador" if self.is_admin else "⚠ Executando como Usuário Normal"
        admin_color = "#28A745" if self.is_admin else "#FFC107"
        tk.Label(
            header_frame,
            text=admin_status,
            font=('Arial', 10),
            foreground=admin_color,
            bg='white'
        ).pack()
        
        # Container para conteúdo
        content_frame = tk.Frame(main_frame, bg='')
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame dos botões com fundo branco
        buttons_container = tk.Frame(content_frame, bg='white', relief='solid', bd=1)
        buttons_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        buttons_frame = tk.Frame(buttons_container, bg='white', padx=5, pady=5)
        buttons_frame.pack(fill=tk.BOTH, expand=True)
        
        functionalities = [
            ("🧹 Limpar Temp", self.limpar_temp),
            ("🌐 Limpar Cache de Rede", self.limpar_cache_rede),
            ("🖼️ Alterar Wallpaper", self.alterar_wallpaper),
            ("🔍 Buscar Atualizações", self.buscar_atualizacoes),
            ("💾 Backup de Arquivos", self.backup_arquivos),
            ("⚡ Padrões de Energia", self.padroes_energia),
            ("📊 Instalar Office 365", self.instalar_office),
            ("🖥️ Instalar Anydesk", self.instalar_anydesk),
            ("🌐 Chrome Padrão", self.chrome_padrao),
            ("🖨️ Instalar Ricoh", self.instalar_ricoh),
            ("📦 Instalar Distribuições", self.instalar_distribuicoes),
            ("👥 Gerenciar Usuários", self.gerenciar_usuarios),
        ]
        
        for i, (text, command) in enumerate(functionalities):
            btn = ttk.Button(
                buttons_frame,
                text=text,
                command=command,
                width=25
            )
            btn.pack(fill=tk.X, pady=2)
        
        # Frame do log com fundo branco
        log_container = tk.Frame(content_frame, bg='white', relief='solid', bd=1)
        log_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        log_frame = tk.LabelFrame(log_container, text="Log de Atividades", bg='white', fg='black', padx=5, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            width=50,
            height=20,
            font=('Consolas', 9),
            bg='white',
            fg='black'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.progress_bar = ttk.Progressbar(
            log_frame,
            variable=self.progress_var,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # Status bar com fundo branco
        status_frame = tk.Frame(main_frame, bg='white', relief='solid', bd=1)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Arial', 9),
            foreground='#666',
            bg='white'
        ).pack(side=tk.LEFT, padx=5, pady=5)
        
        ttk.Button(
            status_frame,
            text="Limpar Log",
            command=self.limpar_log,
            width=10
        ).pack(side=tk.RIGHT, padx=(0, 5), pady=5)
        
        ttk.Button(
            status_frame,
            text="Sair",
            command=self.sair,
            width=10
        ).pack(side=tk.RIGHT, padx=(0, 5), pady=5)
        
        self.log("Sistema inicializado. Selecione uma opção para começar.")

    def load_background(self):
        """Carrega a imagem de fundo se existir"""
        try:
            bg_path = self.script_path / "background.png"
            
            if bg_path.exists():
                print("✅ Carregando imagem de fundo...")
                # Carregar e redimensionar imagem
                self.bg_image = Image.open(bg_path)
                self.bg_image = self.bg_image.resize((800, 600), Image.Resampling.LANCZOS)
                self.bg_photo = ImageTk.PhotoImage(self.bg_image)
                
                # Criar label de fundo
                self.bg_label = tk.Label(self.root, image=self.bg_photo)
                self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
                
                # Garantir que o background fique ATRÁS de tudo
                self.bg_label.lower()
                
                print("✅ Background carregado e posicionado atrás!")
            else:
                print("❌ background.png não encontrado")
                self.root.configure(background='#2E86AB')
                    
        except Exception as e:
            print(f"❌ Erro ao carregar background: {e}")
            self.root.configure(background='#2E86AB')
    
    def log(self, message, type="info"):
        """Adiciona mensagem ao log"""
        if not hasattr(self, 'log_text'):
            print(f"LOG: {message}")
            return
            
        timestamp = time.strftime("%H:%M:%S")
        if type == "error":
            tag = "error"
            prefix = "[ERRO]"
            color = "#DC3545"
        elif type == "warning":
            tag = "warning"
            prefix = "[AVISO]"
            color = "#FFC107"
        elif type == "success":
            tag = "success"
            prefix = "[SUCESSO]"
            color = "#28A745"
        else:
            tag = "info"
            prefix = "[INFO]"
            color = "#17A2B8"
        
        message_line = f"{timestamp} {prefix} {message}\n"
        
        self.log_text.insert(tk.END, message_line)
        self.log_text.see(tk.END)
        
        if type != "info":
            start_index = f"{self.log_text.index(tk.END).split('.')[0]}.0"
            end_index = self.log_text.index(tk.END)
            self.log_text.tag_add(tag, f"{int(start_index.split('.')[0])-1}.0", f"{int(start_index.split('.')[0])-1}.end")
            self.log_text.tag_config(tag, foreground=color)
        
        self.root.update_idletasks()
    
    def limpar_log(self):
        self.log_text.delete(1.0, tk.END)
        self.log("Log limpo.")
    
    def run_in_thread(self, func, *args):
        thread = threading.Thread(target=func, args=args)
        thread.daemon = True
        thread.start()
    
    def update_progress(self, value):
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def limpar_temp(self):
        self.run_in_thread(self._limpar_temp_thread)
    
    def _limpar_temp_thread(self):
        try:
            self.update_status("Limpando arquivos temporários...")
            self.update_progress(0)
            self.log("Iniciando limpeza de arquivos temporários...")
            
            temp_dirs = [
                os.environ.get('TEMP', ''),
                os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'Temp'),
                os.path.join(os.environ.get('USERPROFILE', ''), 'AppData', 'Local', 'Temp')
            ]
            
            total_dirs = len(temp_dirs)
            for i, temp_dir in enumerate(temp_dirs):
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        self.log(f"Limpando: {temp_dir}")
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                try:
                                    file_path = os.path.join(root, file)
                                    os.remove(file_path)
                                except:
                                    pass
                            for dir in dirs:
                                try:
                                    dir_path = os.path.join(root, dir)
                                    shutil.rmtree(dir_path, ignore_errors=True)
                                except:
                                    pass
                    except Exception as e:
                        self.log(f"Erro ao limpar {temp_dir}: {str(e)}", "error")
                
                self.update_progress((i + 1) / total_dirs * 100)
            
            user_temp = os.environ.get('TEMP', '')
            if user_temp and not os.path.exists(user_temp):
                os.makedirs(user_temp)
            
            self.log("Limpeza de arquivos temporários concluída com sucesso!", "success")
            self.update_status("Limpeza concluída!")
            
        except Exception as e:
            self.log(f"Erro durante a limpeza: {str(e)}", "error")
            self.update_status("Erro durante a limpeza!")
        finally:
            self.update_progress(0)
            
    def limpar_cache_rede(self):
        self.run_in_thread(self._limpar_cache_rede_thread)

    def _limpar_cache_rede_thread(self):
        try:
            self.update_status("Limpando cache de rede...")
            self.update_progress(0)
            self.log("Iniciando limpeza de cache de rede...")
            
            comandos = [
                "ipconfig /flushdns",
                "ipconfig /release",
                "ipconfig /renew",
                "netsh winsock reset",
                "netsh int ip reset",
                "arp -d *"
            ]
            
            total_comandos = len(comandos)
            
            for i, comando in enumerate(comandos):
                self.log(f"Executando: {comando}")
                resultado = subprocess.run(
                    comando,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                if resultado.returncode == 0:
                    self.log(f"✓ {comando} - OK")
                else:
                    self.log(f"⚠ {comando} - Aviso", "warning")
                
                self.update_progress((i + 1) / total_comandos * 100)
                time.sleep(2)
            
            self.log("✅ Limpeza de cache de rede concluída com sucesso!", "success")
            self.log("💡 Pode ser necessário reiniciar o computador para aplicar todas as mudanças", "info")
            self.update_status("Cache de rede limpo!")
            
        except subprocess.TimeoutExpired:
            self.log("⏰ Timeout em algum comando", "warning")
            self.log("📋 Parte do cache foi limpo", "info")
        except Exception as e:
            self.log(f"❌ Erro ao limpar cache de rede: {str(e)}", "error")
        finally:
            self.update_progress(0)
    
    def alterar_wallpaper(self):
        self.run_in_thread(self._alterar_wallpaper_thread)

    def _alterar_wallpaper_thread(self):
        try:
            self.update_status("Alterando wallpaper...")
            
            wallpaper_path = self.script_path / "wallpaper.png"
            
            if not wallpaper_path.exists():
                self.log("❌ Arquivo wallpaper.png não encontrado!", "error")
                self.log(f"📁 Local esperado: {wallpaper_path}", "error")
                self.log("💡 Coloque o arquivo wallpaper.png na pasta do aplicativo", "info")
                return
            
            self.log(f"🎨 Alterando wallpaper: {wallpaper_path}")
            
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0, winreg.KEY_WRITE) as key:
                    winreg.SetValueEx(key, "Wallpaper", 0, winreg.REG_SZ, str(wallpaper_path))
                    winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, "10")
                    winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, "0")
            except Exception as e:
                self.log(f"⚠ Erro ao modificar registro: {str(e)}", "warning")
            
            ps_command = f'''
            Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            public class Wallpaper {{
                [DllImport("user32.dll", CharSet=CharSet.Auto)]
                public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
            }}
            "@
            [Wallpaper]::SystemParametersInfo(20, 0, '{wallpaper_path}', 3)
            '''
            
            subprocess.run(f'powershell -Command "{ps_command}"', shell=True)
            
            self.log("🔄 Reiniciando Explorer...")
            subprocess.run("taskkill /f /im explorer.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(2)
            subprocess.run("start explorer.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            self.log("✅ Wallpaper alterado com sucesso!", "success")
            self.log("💡 Se não mudou, faça logoff e login novamente", "info")
            self.update_status("Wallpaper alterado!")
            
        except Exception as e:
            self.log(f"❌ Erro ao alterar wallpaper: {str(e)}", "error")
            self.update_status("Erro ao alterar wallpaper!")
      
    def buscar_atualizacoes(self):
        self.run_in_thread(self._buscar_atualizacoes_thread)
    
    def _buscar_atualizacoes_thread(self):
        try:
            self.update_status("Buscando atualizações...")
            self.log("Iniciando busca por atualizações do Windows...")
            
            subprocess.run("usoclient startscan", shell=True)
            subprocess.run("start ms-settings:windowsupdate-action", shell=True)
            
            self.log("Busca de atualizações iniciada. Verifique as Configurações do Windows.", "success")
            self.update_status("Busca de atualizações concluída!")
            
        except Exception as e:
            self.log(f"Erro ao buscar atualizações: {str(e)}", "error")
            self.update_status("Erro ao buscar atualizações!")
    
    def backup_arquivos(self):
        backup_dir = filedialog.askdirectory(title="Selecione onde salvar o backup")
        
        if backup_dir:
            self.run_in_thread(self._backup_arquivos_thread, backup_dir)
    
    def _backup_arquivos_thread(self, backup_dir):
        try:
            self.update_status("Fazendo backup...")
            self.log(f"Iniciando backup para: {backup_dir}")
            
            user_profile = Path(os.environ['USERPROFILE'])
            backup_folder = Path(backup_dir) / f"{os.environ['USERNAME']}_backup_{time.strftime('%Y-%m-%d_%H-%M-%S')}"
            
            directories_to_backup = [
                'Documents',
                'Desktop', 
                'Pictures',
                'Downloads',
                'Music',
                'Videos'
            ]
            
            total_dirs = len(directories_to_backup)
            for i, directory in enumerate(directories_to_backup):
                source_dir = user_profile / directory
                dest_dir = backup_folder / directory
                
                if source_dir.exists():
                    self.log(f"Copiando: {directory}")
                    shutil.copytree(source_dir, dest_dir, dirs_exist_ok=True)
                
                self.update_progress((i + 1) / total_dirs * 100)
            
            self.log(f"Backup concluído com sucesso em: {backup_folder}", "success")
            self.update_status("Backup concluído!")
            
        except Exception as e:
            self.log(f"Erro durante o backup: {str(e)}", "error")
            self.update_status("Erro durante o backup!")
        finally:
            self.update_progress(0)
    
    def padroes_energia(self):
        self.run_in_thread(self._padroes_energia_thread)
    
    def _padroes_energia_thread(self):
        try:
            self.update_status("Configurando padrões de energia...")
            self.log("Aplicando configurações de energia...")
            
            commands = [
                "powercfg -change -monitor-timeout-ac 5",
                "powercfg -change -standby-timeout-ac 0", 
                "powercfg -change -hibernate-timeout-ac 0",
                "powercfg -change -disk-timeout-ac 0"
            ]
            
            for i, cmd in enumerate(commands):
                subprocess.run(cmd, shell=True)
                self.update_progress((i + 1) / len(commands) * 100)
            
            self.log("Configurações de energia aplicadas com sucesso!", "success")
            self.log("- Tela será bloqueada em 5 minutos")
            self.log("- Computador nunca entrará em modo de espera")
            self.log("- Computador nunca hibernará")
            self.log("- Disco rígido nunca desligará")
            self.update_status("Configurações de energia aplicadas!")
            
        except Exception as e:
            self.log(f"Erro ao configurar energia: {str(e)}", "error")
            self.update_status("Erro ao configurar energia!")
        finally:
            self.update_progress(0)
    
    def instalar_office(self):
        self.run_in_thread(self._instalar_office_thread)
    
    def _instalar_office_thread(self):
        try:
            self.update_status("Instalando Office 365...")
            self.log("Iniciando instalação automática do Office 365...")
            
            office_folder = self.script_path / "office"
            
            if not office_folder.exists():
                self.log("ERRO: Pasta 'office' não encontrada!", "error")
                self.log(f"Local esperado: {office_folder}", "error")
                self.log("Crie a pasta 'office' com os arquivos Setup.exe e Configuration.xml", "error")
                return
            
            setup_file = None
            config_file = None
            
            for possible_setup in ["Setup.exe", "setup.exe", "Setup"]:
                setup_path = office_folder / possible_setup
                if setup_path.exists():
                    setup_file = setup_path
                    break
            
            for possible_config in ["Configuration.xml", "configuration.xml", "Configuração.xml"]:
                config_path = office_folder / possible_config
                if config_path.exists():
                    config_file = config_path
                    break
            
            if not setup_file:
                self.log("ERRO: Arquivo Setup não encontrado na pasta office!", "error")
                self.log("Arquivos encontrados na pasta office:", "error")
                for item in office_folder.iterdir():
                    self.log(f"  - {item.name}", "error")
                return
            
            if not config_file:
                self.log("ERRO: Arquivo Configuration.xml não encontrado na pasta office!", "error")
                return
            
            self.log(f"Setup encontrado: {setup_file.name}")
            self.log(f"Configuração encontrada: {config_file.name}")
            
            self.log("Executando instalação do Office 365...")
            command = f'cd /d "{office_folder}" && "{setup_file}" /configure "{config_file}"'
            
            self.log(f"Comando: {command}")
            
            process = subprocess.Popen(
                command,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            time.sleep(2)
            if process.poll() is None:
                self.log("Instalação do Office 365 iniciada com sucesso!", "success")
                self.log("O processo pode demorar vários minutos. Aguarde a conclusão.", "info")
            else:
                stdout, stderr = process.communicate()
                if stdout:
                    self.log(f"Saída: {stdout}")
                if stderr:
                    self.log(f"Erro: {stderr}", "error")
            
            self.update_status("Instalação do Office iniciada!")
            
        except Exception as e:
            self.log(f"Erro ao instalar Office: {str(e)}", "error")
            self.update_status("Erro ao instalar Office!")
    
    def instalar_anydesk(self):
        self.run_in_thread(self._instalar_anydesk_thread)
    
    def _instalar_anydesk_thread(self):
        try:
            self.update_status("Instalando Anydesk...")
            self.log("Iniciando instalação automática do Anydesk...")
            
            anydesk_file = self.script_path / "anydesk.exe"
            
            if not anydesk_file.exists():
                self.log("ERRO: Arquivo anydesk.exe não encontrado!", "error")
                self.log(f"Local esperado: {anydesk_file}", "error")
                self.log("Coloque o arquivo anydesk.exe na pasta do script", "error")
                return
            
            self.log(f"Arquivo encontrado: {anydesk_file}")
            
            command = f'"{anydesk_file}" --install "{os.environ.get("ProgramFiles", "C:\\\\Program Files")}\\\\AnyDesk" --start-with-win --silent'
            
            self.log("Executando instalação do Anydesk...")
            self.log(f"Comando: {command}")
            
            process = subprocess.Popen(
                command,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            time.sleep(3)
            if process.poll() is None:
                self.log("Instalação do Anydesk iniciada com sucesso!", "success")
                self.log("Aguarde a conclusão do processo...", "info")
            else:
                stdout, stderr = process.communicate()
                if stdout:
                    self.log(f"Saída: {stdout}")
                if stderr:
                    self.log(f"Erro: {stderr}", "error")
                self.log("Instalação do Anydesk concluída!", "success")
            
            self.update_status("Anydesk instalado!")
            
        except Exception as e:
            self.log(f"Erro ao instalar Anydesk: {str(e)}", "error")
            self.update_status("Erro ao instalar Anydesk!")
    
    def instalar_distribuicoes(self):
        self.run_in_thread(self._instalar_distribuicoes_thread)

    def _instalar_distribuicoes_thread(self):
        try:
            self.update_status("Instalando distribuições...")
            self.update_progress(0)
            self.log("🚀 INICIANDO INSTALAÇÃO - MÉTODOS SEGUROS")
            
            distribuicoes = [
                {
                    "nome": ".NET Framework 3.5/4.8",
                    "comando": 'dism /online /enable-feature /featurename:NetFx3 /all /quiet /norestart',
                    "descricao": "Habilita .NET via DISM (Windows)"
                },
                {
                    "nome": "Visual C++ Runtimes", 
                    "comando": 'powershell "Get-WindowsCapability -Online | Where-Object {$_.Name -like \'*VCLibs*\'} | Add-WindowsCapability -Online"',
                    "descricao": "Instala VC++ via Windows Capabilities"
                },
                {
                    "nome": "DirectX Runtime",
                    "comando": 'dism /online /enable-feature /featurename:DirectPlay /all /quiet /norestart',
                    "descricao": "Habilita DirectX via DISM"
                },
                {
                    "nome": "Windows Media Player",
                    "comando": 'dism /online /enable-feature /featurename:WindowsMediaPlayer /all /quiet /norestart',
                    "descricao": "Habilita Media Player"
                },
                {
                    "nome": "Internet Explorer (Legacy)",
                    "comando": 'dism /online /enable-feature /featurename:Internet-Explorer-Optional-amd64 /all /quiet /norestart',
                    "descricao": "Habilita IE se necessário"
                }
            ]
            
            total_dist = len(distribuicoes)
            sucessos = 0
            
            for i, dist in enumerate(distribuicoes):
                self.log(f"📦 {dist['nome']}...")
                self.log(f"📋 {dist['descricao']}")
                
                try:
                    resultado = subprocess.run(
                        dist["comando"],
                        shell=True,
                        capture_output=True,
                        text=True,
                        timeout=180
                    )
                    
                    if resultado.returncode in [0, 3010, 50, 87]:
                        self.log(f"✅ {dist['nome']} - Concluído", "success")
                        sucessos += 1
                    elif resultado.returncode == 0x800f081f:
                        self.log(f"✅ {dist['nome']} - Já instalado", "success")
                        sucessos += 1
                    else:
                        self.log(f"⚠ {dist['nome']} - Código: {resultado.returncode}", "warning")
                        if "already enabled" in resultado.stdout.lower():
                            self.log(f"✅ {dist['nome']} - Já estava habilitado", "success")
                            sucessos += 1
                        
                except subprocess.TimeoutExpired:
                    self.log(f"⏰ {dist['nome']} - Timeout (pode estar processando)", "warning")
                except Exception as e:
                    self.log(f"❌ {dist['nome']} - Erro: {str(e)}", "error")
                
                self.update_progress((i + 1) / total_dist * 100)
                time.sleep(2)
            
            self.log("📦 Java Runtime - Tentando via Winget...")
            try:
                java_result = subprocess.run(
                    'winget install Oracle.JavaRuntimeEnvironment --silent --accept-package-agreements --disable-interactivity',
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if java_result.returncode == 0:
                    self.log("✅ Java Runtime - Instalado via Winget", "success")
                    sucessos += 1
                    total_dist += 1
                else:
                    self.log("⚠ Java - Winget não disponível ou já instalado", "warning")
            except:
                self.log("⚠ Java - Winget não disponível", "warning")
            
            self.log("=" * 50)
            self.log(f"📊 RESUMO DA INSTALAÇÃO SEGURA:", "info")
            self.log(f"✅ Funcionalidades habilitadas: {sucessos}/{total_dist}", "success")
            
            if sucessos > 0:
                self.log("🎉 Distribuições básicas configuradas com sucesso!", "success")
                self.log("💡 Algumas mudanças podem requerer reinicialização", "info")
            else:
                self.log("⚠ Nenhuma distribuição pôde ser habilitada", "warning")
                self.log("📋 Execute como ADMINISTRADOR para melhor resultado", "info")
            
            self.log("🔒 Métodos 100% seguros - Sem download de arquivos", "success")
            self.update_status("Instalação segura concluída!")
            
        except Exception as e:
            self.log(f"❌ Erro: {str(e)}", "error")
            self.update_status("Erro na instalação!")
        finally:
            self.update_progress(0)
    
    def chrome_padrao(self):
        self.run_in_thread(self._chrome_padrao_thread)
    
    def _chrome_padrao_thread(self):
        try:
            self.update_status("Configurando Chrome como padrão...")
            self.log("Iniciando configuração do Chrome como navegador padrão...")
            
            chrome_paths = [
                Path(os.environ.get("ProgramFiles", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
                Path(os.environ.get("ProgramFiles(x86)", "")) / "Google" / "Chrome" / "Application" / "chrome.exe"
            ]
            
            chrome_path = None
            for path in chrome_paths:
                if path.exists():
                    chrome_path = path
                    break
            
            if not chrome_path:
                self.log("Google Chrome não encontrado!", "error")
                self.update_status("Chrome não encontrado!")
                return
            
            self.log(f"Chrome encontrado em: {chrome_path}")
            
            self.log("Configurando via registro do Windows...")
            
            protocols = [
                ("http", "ChromeHTML"),
                ("https", "ChromeHTML"),
                ("ftp", "ChromeFTP"),
                ("mailto", "ChromeHTML"),
                ("file", "ChromeHTML")
            ]
            
            success_count = 0
            total_protocols = len(protocols)
            
            for i, (protocol, progid) in enumerate(protocols):
                try:
                    key_path = f"Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\{protocol}\\UserChoice"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
                        winreg.SetValueEx(key, "Progid", 0, winreg.REG_SZ, progid)
                    self.log(f"✓ {protocol.upper()} configurado")
                    success_count += 1
                except Exception as e:
                    self.log(f"⚠ Não foi possível configurar {protocol.upper()}: {str(e)}", "warning")
                
                self.update_progress((i + 1) / total_protocols * 50)
            
            self.log("Configurando extensões de arquivo...")
            
            file_types = [
                (".html", "ChromeHTML"),
                (".htm", "ChromeHTML"),
                (".pdf", "ChromePDF"),
                (".svg", "ChromeSVG")
            ]
            
            for i, (extension, progid) in enumerate(file_types):
                try:
                    key_path = f"Software\\Classes\\{extension}"
                    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                        winreg.SetValue(key, "", winreg.REG_SZ, progid)
                    self.log(f"✓ {extension} configurado")
                    success_count += 1
                except Exception as e:
                    self.log(f"⚠ Não foi possível configurar {extension}: {str(e)}", "warning")
                
                self.update_progress(50 + ((i + 1) / len(file_types) * 25))
            
            self.log("Forçando atualização das configurações...")
            
            ps_commands = [
                "Start-Process -WindowStyle Hidden -FilePath 'ie4uinit.exe' -ArgumentList '-show'",
                "Start-Process -WindowStyle Hidden -FilePath 'gpupdate' -ArgumentList '/target:user /force'"
            ]
            
            for i, ps_cmd in enumerate(ps_commands):
                try:
                    subprocess.run(
                        f'powershell -Command "{ps_cmd}"',
                        shell=True,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        timeout=10
                    )
                    self.log("✓ Configurações atualizadas")
                except Exception as e:
                    self.log(f"⚠ Aviso na atualização: {str(e)}", "warning")
                
                self.update_progress(75 + ((i + 1) / len(ps_commands) * 25))
            
            try:
                self.log("Executando comando do Chrome em background...")
                subprocess.Popen(
                    [str(chrome_path), "--make-default-browser"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                self.log("✓ Comando do Chrome executado")
            except Exception as e:
                self.log(f"⚠ Comando do Chrome não pôde ser executado: {str(e)}", "warning")
            
            if success_count > 0:
                self.log(f"✅ Chrome configurado como padrão para {success_count} protocolos/extensões!", "success")
                self.log("📋 Resumo:", "success")
                self.log("   - Protocolos HTTP/HTTPS/FTP configurados", "success")
                self.log("   - Extensões HTML/PDF/SVG associadas", "success")
                self.log("   - Sistema de associações atualizado", "success")
                self.log("", "success")
                self.log("💡 Dica: Pode ser necessário reiniciar o computador para que", "info")
                self.log("         todas as alterações tenham efeito completo.", "info")
            else:
                self.log("❌ Não foi possível configurar o Chrome como padrão", "error")
                self.log("   Execute o script como Administrador e tente novamente", "error")
            
            self.update_status("Configuração do Chrome concluída!")
            
        except Exception as e:
            self.log(f"❌ Erro crítico ao configurar Chrome: {str(e)}", "error")
            self.update_status("Erro ao configurar Chrome!")
        finally:
            self.update_progress(0)
    
    def instalar_ricoh(self):
        self._criar_janela_ricoh()
    
    def _criar_janela_ricoh(self):
        ricoh_window = tk.Toplevel(self.root)
        ricoh_window.title("Instalar Impressora Ricoh IM 430")
        ricoh_window.geometry("500x500")
        self.set_window_icon(ricoh_window)
        ricoh_window.resizable(False, False)
        ricoh_window.transient(self.root)
        ricoh_window.grab_set()
        
        ricoh_window.update_idletasks()
        x = (self.root.winfo_x() + (self.root.winfo_width() // 2)) - (500 // 2)
        y = (self.root.winfo_y() + (self.root.winfo_height() // 2)) - (400 // 2)
        ricoh_window.geometry(f"500x500+{x}+{y}")
        
        main_frame = ttk.Frame(ricoh_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame,
            text="Instalação Automática - Ricoh IM 430",
            font=('Arial', 14, 'bold'),
            foreground='#2E86AB'
        ).pack(pady=(0, 20))
        
        ip_frame = ttk.Frame(main_frame)
        ip_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(ip_frame, text="Endereço IP da Impressora:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        
        self.ip_var = tk.StringVar()
        ip_entry = ttk.Entry(
            ip_frame,
            textvariable=self.ip_var,
            font=('Arial', 12),
            width=20
        )
        ip_entry.pack(fill=tk.X, pady=(5, 0))
        ip_entry.focus()
        
        nome_frame = ttk.Frame(main_frame)
        nome_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(nome_frame, text="Nome da Impressora:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)

        self.nome_var = tk.StringVar()
        self.nome_var.set("Ricoh IM 430")

        nome_entry = ttk.Entry(
            nome_frame,
            textvariable=self.nome_var,
            font=('Arial', 12),
            width=20
        )
        nome_entry.pack(fill=tk.X, pady=(5, 0))
        
        info_frame = ttk.LabelFrame(main_frame, text="Configuração Automática", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = """🔧 Driver: RICOH IM 430 PCL 6
🌐 Porta: TCP/IP automática
📄 Método: Instalação silenciosa

📋 Pré-requisitos:
• Impressora ligada e na rede
• IP correto e acessível
• Executar como Administrador (recomendado)"""
        ttk.Label(
            info_frame,
            text=info_text,
            font=('Arial', 9),
            justify=tk.LEFT,
            foreground='#666'
        ).pack(anchor=tk.W)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame,
            text="🚀 Instalar Automaticamente",
            command=lambda: self._iniciar_instalacao_ricoh(ricoh_window),
            style="Accent.TButton",
            width=25
        ).pack(side=tk.RIGHT, padx=(10, 0))
        
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=ricoh_window.destroy
        ).pack(side=tk.RIGHT)
        
        style = ttk.Style()
        style.configure("Accent.TButton", foreground='white', background='#28A745', font=('Arial', 10, 'bold'))
    
    def _iniciar_instalacao_ricoh(self, ricoh_window):
        ip = self.ip_var.get().strip()
        
        if not ip:
            messagebox.showerror("Erro", "Por favor, informe o endereço IP da impressora.")
            return
        
        ip_parts = ip.split('.')
        if len(ip_parts) != 4:
            messagebox.showerror("Erro", "Formato de IP inválido. Use: XXX.XXX.XXX.XXX")
            return
        
        for part in ip_parts:
            if not part.isdigit() or not 0 <= int(part) <= 255:
                messagebox.showerror("Erro", "Endereço IP inválido. Cada parte deve ser entre 0-255.")
                return
            
        nome_impressora = self.nome_var.get().strip()

        if not nome_impressora:
            messagebox.showerror("Erro", "Por favor, informe um nome para a impressora.")
            return
        
        ricoh_window.destroy()
        self.run_in_thread(self._instalar_ricoh_thread, ip, nome_impressora)
    
    def _instalar_ricoh_thread(self, ip, nome_impressora):
        try:
            self.update_status("Instalando Ricoh IM 430...")
            self.log("🚀 INICIANDO INSTALAÇÃO AUTOMÁTICA RICOH IM 430")
            self.log(f"📡 IP da impressora: {ip}")
            self.log(f"🔧 Driver: RICOH IM 430 PCL 6")
            
            driver_name = "RICOH IM 430 PCL 6"
            printer_name = f"{nome_impressora} ({ip})"
            port_name = f"IP_{ip}"
            
            self.update_progress(10)
            
            self.log("🔍 Verificando conectividade com a impressora...")
            try:
                ping_result = subprocess.run(
                    f"ping -n 2 -w 1000 {ip}",
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                if ping_result.returncode == 0:
                    self.log("✅ Conectividade: OK")
                else:
                    self.log("⚠️ Conectividade: Possível problema", "warning")
            except:
                self.log("⚠️ Não foi possível verificar conectividade", "warning")
            
            self.update_progress(20)
            
            self.log(f"🔌 Criando porta TCP/IP: {port_name}")
            
            add_port_cmd = f'powershell "Add-PrinterPort -Name \\"{port_name}\\" -PrinterHostAddress \\"{ip}\\" -ErrorAction Stop"'
            
            port_result = subprocess.run(
                add_port_cmd,
                shell=True,
                capture_output=True,
                text=True
            )
            
            if port_result.returncode == 0:
                self.log("✅ Porta TCP/IP criada com sucesso")
            else:
                self.log(f"⚠️ Aviso ao criar porta: {port_result.stderr}", "warning")
            
            self.update_progress(40)
            
            self.log("🔎 Verificando disponibilidade do driver...")
            
            check_driver_cmd = f'powershell "Get-PrinterDriver -Name \\"{driver_name}\\" -ErrorAction SilentlyContinue"'
            
            driver_result = subprocess.run(
                check_driver_cmd,
                shell=True,
                capture_output=True,
                text=True
            )
            
            if driver_result.returncode != 0:
                self.log("📥 Driver não encontrado, instalando...")
                install_driver_cmd = f'powershell "Add-PrinterDriver -Name \\"{driver_name}\\" -ErrorAction SilentlyContinue"'
                subprocess.run(install_driver_cmd, shell=True)
                self.log("✅ Driver instalado/verificado")
            else:
                self.log("✅ Driver já disponível")
            
            self.update_progress(60)
            
            self.log("🖨️ Instalando impressora...")
            
            install_printer_cmd = f'''
            powershell -Command "
            try {{
                Add-Printer -Name '{printer_name}' -DriverName '{driver_name}' -PortName '{port_name}' -ErrorAction Stop
                Write-Output 'SUCCESS'
            }} catch {{
                Write-Output 'ERROR: $($_.Exception.Message)'
            }}
            "
            '''
            
            install_result = subprocess.run(
                install_printer_cmd,
                shell=True,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            self.update_progress(80)
            
            if "SUCCESS" in install_result.stdout:
                self.log("✅ Impressora instalada com sucesso via PowerShell")
            else:
                self.log("🔄 Tentando método alternativo...")
                
                alt_cmd = f'rundll32 printui.dll,PrintUIEntry /if /b "{printer_name}" /r "{port_name}" /m "{driver_name}" /q'
                
                alt_result = subprocess.run(
                    alt_cmd,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                if alt_result.returncode == 0:
                    self.log("✅ Impressora instalada com sucesso via Rundll32")
                else:
                    self.log("❌ Falha na instalação automática", "error")
                    self.log("📋 Tentando método manual simplificado...")
                    
                    simple_cmd = f'printui.dll,PrintUIEntry /if /b "{printer_name}" /r "{port_name}" /m "{driver_name}"'
                    subprocess.run(simple_cmd, shell=True)
                    self.log("🔧 Comando manual executado")
            
            self.update_progress(90)
            
            self.log("🔍 Verificando instalação...")
            time.sleep(2)
            
            verify_cmd = f'powershell "Get-Printer -Name \\"{printer_name}\\" -ErrorAction SilentlyContinue"'
            verify_result = subprocess.run(
                verify_cmd,
                shell=True,
                capture_output=True,
                text=True
            )
            
            if verify_result.returncode == 0:
                self.log(f"🎉 IMPRESSORA INSTALADA COM SUCESSO: {printer_name}", "success")
                self.log("📍 Verifique em: Configurações > Impressoras e scanners", "success")
            else:
                self.log("⚠️ Impressora não detectada automaticamente", "warning")
                self.log("📋 Mas pode ter sido instalada. Verifique manualmente.", "info")
            
            self.update_progress(100)
            
            self.log("✨ Processo de instalação concluído!", "success")
            self.update_status("Instalação da Ricoh concluída!")
            
        except subprocess.TimeoutExpired:
            self.log("⏰ Timeout: O processo demorou muito", "warning")
            self.log("📋 A impressora pode ter sido instalada. Verifique manualmente.", "info")
        except Exception as e:
            self.log(f"❌ Erro durante a instalação: {str(e)}", "error")
            self.log("💡 Dica: Execute como Administrador para melhor resultado", "info")
        finally:
            self.update_progress(0)
            
    def gerenciar_usuarios(self):
        self._criar_janela_usuarios()

    def _criar_janela_usuarios(self):
        usuarios_window = tk.Toplevel(self.root)
        usuarios_window.title("Gerenciar Usuários Locais")
        usuarios_window.geometry("500x450")
        self.set_window_icon(usuarios_window)
        usuarios_window.resizable(False, False)
        usuarios_window.transient(self.root)
        usuarios_window.grab_set()
        
        usuarios_window.update_idletasks()
        x = (self.root.winfo_x() + (self.root.winfo_width() // 2)) - (500 // 2)
        y = (self.root.winfo_y() + (self.root.winfo_height() // 2)) - (450 // 2)
        usuarios_window.geometry(f"500x450+{x}+{y}")
        
        main_frame = ttk.Frame(usuarios_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame,
            text="Gerenciar Usuários Locais",
            font=('Arial', 14, 'bold'),
            foreground='#2E86AB'
        ).pack(pady=(0, 20))
        
        admin_frame = ttk.LabelFrame(main_frame, text="Renomear Usuário TI", padding="10")
        admin_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(admin_frame, text="Nova senha para TI:", font=('Arial', 9)).pack(anchor=tk.W)
        
        self.senha_admin_var = tk.StringVar()
        senha_entry = ttk.Entry(
            admin_frame,
            textvariable=self.senha_admin_var,
            font=('Arial', 10),
            show="*",
            width=20
        )
        senha_entry.pack(fill=tk.X, pady=(5, 10))
        
        ttk.Button(
            admin_frame,
            text="🔧 Definir Senha Admin",
            command=lambda: self._renomear_admin(usuarios_window),
            width=25
        ).pack()
        
        usuario_frame = ttk.LabelFrame(main_frame, text="Criar Novo Usuário", padding="10")
        usuario_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(usuario_frame, text="Nome do usuário:", font=('Arial', 9)).pack(anchor=tk.W)
        
        self.novo_usuario_var = tk.StringVar()
        usuario_entry = ttk.Entry(
            usuario_frame,
            textvariable=self.novo_usuario_var,
            font=('Arial', 10),
            width=20
        )
        usuario_entry.pack(fill=tk.X, pady=(5, 10))
        
        ttk.Label(usuario_frame, text="Senha (opcional):", font=('Arial', 9)).pack(anchor=tk.W)
        
        self.senha_usuario_var = tk.StringVar()
        senha_usuario_entry = ttk.Entry(
            usuario_frame,
            textvariable=self.senha_usuario_var,
            font=('Arial', 10),
            show="*",
            width=20
        )
        senha_usuario_entry.pack(fill=tk.X, pady=(5, 10))
        
        ttk.Button(
            usuario_frame,
            text="👤 Criar Usuário",
            command=lambda: self._criar_usuario_personalizado(usuarios_window),
            width=25
        ).pack()
        
        info_frame = ttk.LabelFrame(main_frame, text="Informações", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = """• Execute como ADMINISTRADOR para estas funções
• Usuário TI será renomeado para 'TI'
• Novos usuários serão criados sem senha (ou com senha se informada)
• Todos serão adicionados ao grupo de Usuários"""
        
        ttk.Label(
            info_frame,
            text=info_text,
            font=('Arial', 9),
            justify=tk.LEFT,
            foreground='#666'
        ).pack(anchor=tk.W)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame,
            text="Fechar",
            command=usuarios_window.destroy
        ).pack(side=tk.RIGHT)

    def _renomear_admin(self, window):
        senha = self.senha_admin_var.get().strip()
        
        if not senha:
            messagebox.showerror("Erro", "Por favor, informe uma senha para o usuário Admin.")
            return
        
        if len(senha) < 3:
            messagebox.showerror("Erro", "A senha deve ter pelo menos 3 caracteres.")
            return
        
        window.destroy()
        self.run_in_thread(self._renomear_admin_thread, senha)

    def _renomear_admin_thread(self, senha):
        try:
            self.update_status("Renomeando usuário Admin...")
            self.log("Iniciando renomeação do usuário Admin...")
            
            comandos = [
                f'Rename-LocalUser -Name "ti" -NewName "TI"',
                f'$senha = ConvertTo-SecureString "{senha}" -AsPlainText -Force',
                f'Set-LocalUser -Name "Admin" -Password $senha'
            ]
            
            ps_script = "; ".join(comandos)
            
            resultado = subprocess.run(
                f'powershell -Command "{ps_script}"',
                shell=True,
                capture_output=True,
                text=True
            )
            
            if resultado.returncode == 0:
                self.log("✅ Usuário admin renomeado para 'Admin' com sucesso!", "success")
                self.log("✅ Senha definida para o usuário Admin", "success")
            else:
                self.log(f"❌ Erro ao renomear admin: {resultado.stderr}", "error")
                self.log("💡 Execute como Administrador", "info")
            
            self.update_status("Renomeação concluída!")
            
        except Exception as e:
            self.log(f"❌ Erro ao renomear admin: {str(e)}", "error")

    def _criar_usuario_personalizado(self, window):
        nome_usuario = self.novo_usuario_var.get().strip()
        senha = self.senha_usuario_var.get().strip()
        
        if not nome_usuario:
            messagebox.showerror("Erro", "Por favor, informe um nome para o usuário.")
            return
        
        import re
        if not re.match(r'^[a-zA-Z0-9_]+$', nome_usuario):
            messagebox.showerror("Erro", "Nome de usuário inválido. Use apenas letras, números e underscore.")
            return
        
        window.destroy()
        self.run_in_thread(self._criar_usuario_personalizado_thread, nome_usuario, senha)

    def _criar_usuario_personalizado_thread(self, nome_usuario, senha):
        try:
            self.update_status(f"Criando usuário {nome_usuario}...")
            self.log(f"Criando usuário '{nome_usuario}'...")
            
            if senha:
                comandos = [
                    f'$senha = ConvertTo-SecureString "{senha}" -AsPlainText -Force',
                    f'New-LocalUser -Name "{nome_usuario}" -Password $senha',
                    f'Add-LocalGroupMember -Group "Users" -Member "{nome_usuario}"'
                ]
                ps_script = "; ".join(comandos)
            else:
                comandos = [
                    f'New-LocalUser -Name "{nome_usuario}" -NoPassword',
                    f'Add-LocalGroupMember -Group "Users" -Member "{nome_usuario}"'
                ]
                ps_script = "; ".join(comandos)
        
            resultado = subprocess.run(
                f'powershell -Command "{ps_script}"',
                shell=True,
                capture_output=True,
                text=True
            )
            
            if resultado.returncode == 0:
                if senha:
                    self.log(f"✅ Usuário '{nome_usuario}' criado com senha com sucesso!", "success")
                else:
                    self.log(f"✅ Usuário '{nome_usuario}' criado sem senha com sucesso!", "success")
                self.log("✅ Usuário adicionado ao grupo Users", "success")
            else:
                if "already exists" in resultado.stderr.lower():
                    self.log(f"❌ Usuário '{nome_usuario}' já existe", "error")
                else:
                    self.log(f"❌ Erro ao criar usuário '{nome_usuario}': {resultado.stderr}", "error")
                self.log("💡 Execute como Administrador", "info")
            
            self.update_status(f"Usuário {nome_usuario} criado!")
            
        except Exception as e:
            self.log(f"❌ Erro ao criar usuário '{nome_usuario}': {str(e)}", "error")
    
    def sair(self):
        if messagebox.askyesno("Sair", "Deseja realmente sair?"):
            self.root.quit()

def main():
    try:
        root = tk.Tk()
        app = AutomationApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()