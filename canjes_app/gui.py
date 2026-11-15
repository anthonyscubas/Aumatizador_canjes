"""Aplicación de escritorio Tkinter."""
import ctypes
import datetime as dt
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import keyring
from .config import SERVICE, USER, CORREO_REMITENTE, CARPETA_CANJES
from .exporter import exportar
from .mailer import enviar_correos_outlook

if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()


        self.title('EMPRESA | Exportador de Canjes')
        self.geometry('1180x800')
        self.minsize(950, 700)

        BG = '#0B1220'
        CARD = '#172338'
        INPUT = '#101C30'
        TEXT = '#F8FAFC'
        MUTED = '#94A3B8'
        BLUE = '#2563EB'
        GREEN = '#0F766E'
        BORDER = '#334155'

        self.configure(bg=BG)

        self.events = queue.Queue()
        self.en_proceso = False

        hoy = dt.date.today()

        self.desde = tk.StringVar(
            value=(hoy - dt.timedelta(days=1)).strftime('%d/%m/%Y')
        )
        self.hasta = tk.StringVar(
            value=hoy.strftime('%d/%m/%Y')
        )
        self.carpeta = tk.StringVar(
            value=str(CARPETA_CANJES)
        )

        try:
            guardada = keyring.get_password(SERVICE, USER) or ''
        except Exception:
            guardada = ''

        self.password = tk.StringVar(value=guardada)
        self.recordar = tk.BooleanVar(value=True)
        self.estado = tk.StringVar(value='Sistema listo')

        estilo = ttk.Style(self)
        estilo.theme_use('clam')

        estilo.configure(
            'Modern.TEntry',
            padding=10,
            fieldbackground=INPUT,
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER
        )

        estilo.configure(
            'Modern.TCheckbutton',
            background=CARD,
            foreground=TEXT,
            font=('Segoe UI', 10)
        )
        estilo.map(
            'Modern.TCheckbutton',
            background=[('active', CARD)],
            foreground=[('active', TEXT)]
        )

        estilo.configure(
            'Modern.Horizontal.TProgressbar',
            troughcolor=INPUT,
            background=BLUE,
            borderwidth=0,
            thickness=8
        )

        estilo.configure(
            'Modern.DateEntry',
            font=('Segoe UI', 12),
            fieldbackground='white',
            foreground='#111827',
            padding=6
        )


        root = tk.Frame(self, bg=BG)
        root.pack(fill='both', expand=True, padx=36, pady=24)

        root.columnconfigure(0, weight=1, uniform='panel')
        root.columnconfigure(1, weight=1, uniform='panel')
        root.rowconfigure(5, weight=1)


        header = tk.Frame(root, bg=BG)
        header.grid(
            row=0, column=0, columnspan=2,
            sticky='ew', pady=(0, 24)
        )

        tk.Label(
            header,
            text='EMPRESA TECHNOLOGY',
            font=('Segoe UI', 10, 'bold'),
            fg='#60A5FA',
            bg=BG
        ).pack(anchor='w')

        tk.Label(
            header,
            text='Exportador de Canjes',
            font=('Segoe UI', 27, 'bold'),
            fg=TEXT,
            bg=BG
        ).pack(anchor='w', pady=(3, 2))

        tk.Label(
            header,
            text='Gestión de reportes, validación y distribución por correo',
            font=('Segoe UI', 11),
            fg=MUTED,
            bg=BG
        ).pack(anchor='w')

        card_fecha = tk.Frame(
            root, bg=CARD, padx=24, pady=22,
            highlightbackground=BORDER, highlightthickness=1
        )
        card_fecha.grid(
            row=1, column=0, sticky='nsew',
            padx=(0, 10), pady=(0, 16)
        )

        card_fecha.columnconfigure(0, weight=1)
        card_fecha.columnconfigure(1, weight=1)

        tk.Label(
            card_fecha,
            text='01  PERIODO DE EXTRACCIÓN',
            font=('Segoe UI', 12, 'bold'),
            fg=TEXT, bg=CARD
        ).grid(
            row=0, column=0, columnspan=2,
            sticky='w', pady=(0, 20)
        )

        tk.Label(
            card_fecha,
            text='FECHA DESDE',
            font=('Segoe UI', 10, 'bold'),
            fg=MUTED, bg=CARD
        ).grid(row=1, column=0, sticky='w')

        tk.Label(
            card_fecha,
            text='FECHA HASTA',
            font=('Segoe UI', 10, 'bold'),
            fg=MUTED, bg=CARD
        ).grid(row=1, column=1, sticky='w', padx=(12, 0))

        self.cal_desde = DateEntry(
            card_fecha,
            date_pattern='dd/mm/yyyy',
            locale='es_ES',
            style='Modern.DateEntry',
            background=BLUE,
            foreground='white',
            borderwidth=2
        )
        self.cal_desde.set_date(hoy - dt.timedelta(days=1))
        self.cal_desde.grid(
            row=2, column=0,
            sticky='ew', pady=(9, 0)
        )

        self.cal_hasta = DateEntry(
            card_fecha,
            date_pattern='dd/mm/yyyy',
            locale='es_ES',
            style='Modern.DateEntry',
            background=BLUE,
            foreground='white',
            borderwidth=2
        )
        self.cal_hasta.set_date(hoy)
        self.cal_hasta.grid(
            row=2, column=1,
            sticky='ew', padx=(12, 0),
            pady=(9, 0)
        )

        card_config = tk.Frame(
            root, bg=CARD, padx=24, pady=22,
            highlightbackground=BORDER, highlightthickness=1
        )
        card_config.grid(
            row=1, column=1, sticky='nsew',
            padx=(10, 0), pady=(0, 16)
        )
        card_config.columnconfigure(0, weight=1)

        tk.Label(
            card_config,
            text='02  CONEXIÓN AZURE SQL',
            font=('Segoe UI', 12, 'bold'),
            fg=TEXT, bg=CARD
        ).grid(row=0, column=0, sticky='w', pady=(0, 15))

        tk.Label(
            card_config,
            text='CONTRASEÑA DE ACCESO',
            font=('Segoe UI', 10, 'bold'),
            fg=MUTED, bg=CARD
        ).grid(row=1, column=0, sticky='w')

        ttk.Entry(
            card_config,
            textvariable=self.password,
            show='•',
            style='Modern.TEntry',
            font=('Segoe UI', 12)
        ).grid(
            row=2, column=0,
            sticky='ew', pady=(9, 10)
        )

        ttk.Checkbutton(
            card_config,
            text='Guardar contraseña en Windows',
            variable=self.recordar,
            style='Modern.TCheckbutton'
        ).grid(row=3, column=0, sticky='w')


        card_salida = tk.Frame(
            root, bg=CARD, padx=24, pady=20,
            highlightbackground=BORDER, highlightthickness=1
        )
        card_salida.grid(
            row=2, column=0, columnspan=2,
            sticky='ew', pady=(0, 16)
        )
        card_salida.columnconfigure(0, weight=1)

        tk.Label(
            card_salida,
            text='03  CARPETA DE DESTINO',
            font=('Segoe UI', 12, 'bold'),
            fg=TEXT, bg=CARD
        ).grid(
            row=0, column=0, columnspan=2,
            sticky='w', pady=(0, 12)
        )

        ttk.Entry(
            card_salida,
            textvariable=self.carpeta,
            style='Modern.TEntry',
            font=('Segoe UI', 11)
        ).grid(
            row=1, column=0,
            sticky='ew', padx=(0, 12)
        )

        tk.Button(
            card_salida,
            text='EXAMINAR',
            command=self.examinar,
            font=('Segoe UI', 10, 'bold'),
            bg='#334155',
            fg='white',
            activebackground='#475569',
            activeforeground='white',
            relief='flat',
            cursor='hand2',
            padx=22,
            pady=9,
            bd=0
        ).grid(row=1, column=1, sticky='ew')

        acciones = tk.Frame(root, bg=BG)
        acciones.grid(
            row=3, column=0, columnspan=2,
            sticky='ew', pady=(2, 18)
        )
        acciones.columnconfigure(0, weight=1)
        acciones.columnconfigure(1, weight=1)

        self.boton = tk.Button(
            acciones,
            text='▣  GENERAR LOS 2 EXCEL',
            command=self.iniciar,
            font=('Segoe UI', 12, 'bold'),
            bg=BLUE,
            fg='white',
            activebackground='#1D4ED8',
            activeforeground='white',
            relief='flat',
            bd=0,
            cursor='hand2',
            pady=16
        )
        self.boton.grid(
            row=0, column=0,
            sticky='ew', padx=(0, 10)
        )

        self.boton_correo = tk.Button(
            acciones,
            text='✉  ENVIAR CORREOS',
            command=self.iniciar_envio_correos,
            font=('Segoe UI', 12, 'bold'),
            bg=GREEN,
            fg='white',
            activebackground='#115E59',
            activeforeground='white',
            relief='flat',
            bd=0,
            cursor='hand2',
            pady=16
        )
        self.boton_correo.grid(
            row=0, column=1,
            sticky='ew', padx=(10, 0)
        )

        estado_frame = tk.Frame(root, bg=BG)
        estado_frame.grid(
            row=4, column=0, columnspan=2,
            sticky='ew', pady=(0, 16)
        )

        tk.Label(
            estado_frame,
            text='ESTADO DEL PROCESO',
            font=('Segoe UI', 10, 'bold'),
            fg=MUTED,
            bg=BG
        ).pack(anchor='w')

        tk.Label(
            estado_frame,
            textvariable=self.estado,
            font=('Segoe UI', 11),
            fg='#86EFAC',
            bg=BG,
            anchor='w',
            wraplength=1050
        ).pack(fill='x', pady=(5, 10))

        self.barra = ttk.Progressbar(
            estado_frame,
            mode='indeterminate',
            style='Modern.Horizontal.TProgressbar'
        )
        self.barra.pack(fill='x')

        consola = tk.Frame(
            root, bg=CARD, padx=20, pady=16,
            highlightbackground=BORDER, highlightthickness=1
        )
        consola.grid(
            row=5, column=0, columnspan=2,
            sticky='nsew'
        )
        consola.columnconfigure(0, weight=1)
        consola.rowconfigure(1, weight=1)

        tk.Label(
            consola,
            text='REGISTRO DE ACTIVIDAD',
            font=('Segoe UI', 11, 'bold'),
            fg=TEXT, bg=CARD
        ).grid(
            row=0, column=0,
            sticky='w', pady=(0, 12)
        )

        log_frame = tk.Frame(consola, bg=INPUT)
        log_frame.grid(row=1, column=0, sticky='nsew')
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(
            log_frame,
            bg=INPUT,
            fg='#86EFAC',
            insertbackground='white',
            font=('Consolas', 11),
            state='disabled',
            relief='flat',
            bd=0,
            padx=14,
            pady=12,
            wrap='word',
            height=7
        )
        self.log.grid(
            row=0, column=0,
            sticky='nsew'
        )

        scroll = tk.Scrollbar(
            log_frame,
            command=self.log.yview,
            bg=CARD,
            troughcolor=INPUT
        )
        scroll.grid(row=0, column=1, sticky='ns')
        self.log.configure(yscrollcommand=scroll.set)

        tk.Label(
            root,
            text='EMPRESA | EXPORTADOR',
            font=('Segoe UI', 9),
            fg='#64748B',
            bg=BG
        ).grid(
            row=6, column=0, columnspan=2,
            pady=(15, 0)
        )

        self.after(200, self.revisar)


    def examinar(self):
        path = filedialog.askdirectory(initialdir=self.carpeta.get())
        if path:
            self.carpeta.set(path)

    def escribir_log(self, mensaje):
        self.estado.set(mensaje)
        self.log.configure(state='normal')
        self.log.insert('end', f'[{dt.datetime.now():%H:%M:%S}] {mensaje}\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def iniciar(self):
        if self.en_proceso:
            return
        try:
            inicio = self.cal_desde.get_date()
            fin = self.cal_hasta.get_date()
            if inicio > fin:
                raise ValueError('Fecha inicial posterior a fecha final')
            if not self.password.get():
                raise ValueError('Debes ingresar contraseña SQL')
        except ValueError as exc:
            messagebox.showwarning('Revisar campos', str(exc))
            return
        try:
            if self.recordar.get():
                keyring.set_password(SERVICE, USER, self.password.get())
            else:
                try:
                    keyring.delete_password(SERVICE, USER)
                except Exception:
                    pass
        except Exception as exc:
            messagebox.showwarning('Credenciales', f'No se pudo actualizar el guardado: {exc}')
        self.en_proceso = True
        self.boton.configure(state='disabled')
        self.boton_correo.configure(state='disabled')
        self.barra.start(12)
        self.escribir_log('INICIANDO EXPORTACIÓN...')
        self.ultima_actividad = time.monotonic()
        self.etapa = 'Preparando extracción'
        self.after(15000, self.latido)
        threading.Thread(target=self.trabajar, args=(inicio, fin, self.carpeta.get(), self.password.get()), daemon=True).start()

    def trabajar(self, inicio, fin, carpeta, clave):
        def aviso(texto):
            self.events.put(('log', texto))
        try:
            resultado = exportar(inicio, fin, carpeta, clave, aviso)
            self.events.put(('ok', resultado))
        except Exception as exc:
            self.events.put(('error', str(exc)))

    def iniciar_envio_correos(self):

        if self.en_proceso:
            return

        try:
            lima, provincia = obtener_ultimos_excels()
            destinatarios = leer_destinatarios()

        except Exception as exc:
            messagebox.showerror(
                'Error al preparar correos',
                str(exc)
            )
            return

        para_lima = len(destinatarios['lima']['para'])
        cc_lima = len(destinatarios['lima']['cc'])

        para_prov = len(destinatarios['provincia']['para'])
        cc_prov = len(destinatarios['provincia']['cc'])

        mensaje = (
            '¿Deseas enviar los siguientes 2 correos?\n\n'

            'REMITENTE:\n'
            f'{CORREO_REMITENTE}\n\n'

            'LIMA - AYACUCHO\n'
            f'Para: {para_lima} destinatarios\n'
            f'CC: {cc_lima} destinatarios\n'
            f'Adjunto: {lima.name}\n\n'

            'PROVINCIA\n'
            f'Para: {para_prov} destinatarios\n'
            f'CC: {cc_prov} destinatarios\n'
            f'Adjunto: {provincia.name}\n\n'

            'Se enviarán únicamente 2 correos.\n\n'
            '¿Confirmar envío?'
        )

        confirmar = messagebox.askyesno(
            'Confirmar envío de correos',
            mensaje
        )

        if not confirmar:
            return

        self.en_proceso = True

        self.boton.configure(state='disabled')
        self.boton_correo.configure(state='disabled')

        self.barra.start(12)

        self.escribir_log(
            'INICIANDO ENVÍO DE CORREOS...'
        )

        self.ultima_actividad = time.monotonic()
        self.etapa = 'Preparando correos'

        self.after(15000, self.latido)

        threading.Thread(
            target=self.trabajar_correos,
            daemon=True
        ).start()


    def trabajar_correos(self):

        def aviso(mensaje):
            self.events.put(('log', mensaje))

        try:
            cantidad = enviar_correos_outlook(aviso)

            self.events.put((
                'correo_ok',
                cantidad
            ))

        except Exception as exc:

            self.events.put((
                'correo_error',
                str(exc)
            ))
    def latido(self):
        if not self.en_proceso:
            return
        segundos = int(time.monotonic() - self.ultima_actividad)
        if segundos >= 15:
            self.escribir_log(f'Esperando {segundos}s en: {self.etapa}')
        self.after(15000, self.latido)

    def revisar(self):
        try:
            while True:
                tipo, dato = self.events.get_nowait()
                if tipo == 'log':
                    self.etapa = dato
                    self.ultima_actividad = time.monotonic()
                    self.escribir_log(dato)
                else:
                    self.en_proceso = False
                    self.barra.stop()
                    self.boton.configure(state='normal')
                    self.boton_correo.configure(state='normal')
                    if tipo == 'ok':

                        rutas, lima, provincia = dato

                        self.escribir_log(
                            f'FINALIZADO: Lima {lima:,} | '
                            f'Provincia {provincia:,}'
                        )

                        messagebox.showinfo(
                            'Exportación terminada',
                            '\n'.join(str(p) for p in rutas)
                        )

                    elif tipo == 'correo_ok':

                        self.escribir_log(
                            f'ENVÍO FINALIZADO: '
                            f'{dato} correos entregados a Outlook.'
                        )

                        messagebox.showinfo(
                            'Correos procesados',
                            f'Outlook aceptó {dato} correos '
                            'para su envío.\n\n'
                            '1. Lima - Ayacucho\n'
                            '2. Provincia\n\n'
                            'Puedes verificar su estado '
                            'en Outlook.'
                        )

                    elif tipo == 'correo_error':

                        self.escribir_log(
                            'ERROR AL ENVIAR: ' + dato
                        )

                        messagebox.showerror(
                            'Error de envío',
                            dato
                        )

                    else:

                        self.escribir_log(
                            'ERROR: ' + dato
                        )

                        messagebox.showerror(
                            'Error',
                            dato
                        )
        except queue.Empty:
            pass
        self.after(200, self.revisar)


