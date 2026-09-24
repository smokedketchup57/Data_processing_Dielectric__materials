import warnings
warnings.filterwarnings("ignore")

import os 
import glob 
import pandas as pd 
import numpy as np 
import io
import sys
import matplotlib.pyplot as plt

EPSILON_0 = 8.854e-14  # Está en centímetros 
pi = np.pi 

# ==============================================================================
# CONFIGURACIÓN MULTI-BIOMASA (Rutas Dinámicas de Entrada y Salida)
# ==============================================================================
Configuracion_Biomasas = [
    {
        "Nombre": "Fruta Milagrosa",
        "Ruta_Base": r"./Fruta_Milagrosa",
        "Ruta_Imagenes": r"./Fruta_Milagrosa/Imagenes",
        "Nombre_Excel": "Reporte_Consolidado_Fruta_Milagrosa.xlsx"
    },
    {
        "Nombre": "Pergamino de Café",
        "Ruta_Base": r"./Coffee_Parchment",
        "Ruta_Imagenes": r"./Coffee_Parchment/Imagenes_PC",
        "Nombre_Excel": "Reporte_Consolidado_Coffee_Parchment.xlsx"
    }
]


# --- TUS DATOS BASE ---
datos_base = [
    {"Biomasa": "Fruta Milagrosa", "Tipo": "SFM", "Muestra": "1er", "Diametro_cm": 1.203, "Distancia_cm": 0.300},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "SFM", "Muestra": "2do", "Diametro_cm": 1.245, "Distancia_cm": 0.285},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "PFM", "Muestra": "1er", "Diametro_cm": 1.240, "Distancia_cm": 0.250},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "PFM", "Muestra": "2do", "Diametro_cm": 1.235, "Distancia_cm": 0.260},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "CFM", "Muestra": "1er", "Diametro_cm": 1.255, "Distancia_cm": 0.345},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "CFM", "Muestra": "2do", "Diametro_cm": 1.275, "Distancia_cm": 0.330},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "SM",  "Muestra": "1er", "Diametro_cm": 1.260, "Distancia_cm": 0.300},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "SM",  "Muestra": "2do", "Diametro_cm": 1.245, "Distancia_cm": 0.280},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "CM",  "Muestra": "1er", "Diametro_cm": 1.245, "Distancia_cm": 0.280},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "CM",  "Muestra": "2do", "Diametro_cm": 1.250, "Distancia_cm": 0.300},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "HFM", "Muestra": "1er", "Diametro_cm": 1.265, "Distancia_cm": 0.315},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "HFM", "Muestra": "2do", "Diametro_cm": 1.265, "Distancia_cm": 0.315},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "HM",  "Muestra": "1er", "Diametro_cm": 1.250, "Distancia_cm": 0.330},
    {"Biomasa": "Fruta Milagrosa", "Tipo": "HM",  "Muestra": "2do", "Diametro_cm": 1.280, "Distancia_cm": 0.340},
    {"Biomasa": "Pergamino de Café", "Tipo": "CMC", "Muestra": "1er", "Diametro_cm": 1.300, "Distancia_cm": 0.395},
    {"Biomasa": "Pergamino de Café", "Tipo": "CMC", "Muestra": "2do", "Diametro_cm": 1.295, "Distancia_cm": 0.385},
    {"Biomasa": "Pergamino de Café", "Tipo": "CPC", "Muestra": "1er", "Diametro_cm": 1.245, "Distancia_cm": 0.335},
    {"Biomasa": "Pergamino de Café", "Tipo": "CPC", "Muestra": "2do", "Diametro_cm": 1.255, "Distancia_cm": 0.325},
    {"Biomasa": "Pergamino de Café", "Tipo": "PCT", "Muestra": "1er", "Diametro_cm": 1.265, "Distancia_cm": 0.345},
    {"Biomasa": "Pergamino de Café", "Tipo": "PCT", "Muestra": "2do", "Diametro_cm": 1.265, "Distancia_cm": 0.365},
]

# --- TRANSFORMACIÓN A DICCIONARIO DE BÚSQUEDA RÁPIDA ---
Geometria_Pellets = {}
for dato in datos_base:
    tipo = dato["Tipo"].lower()
    muestra = dato["Muestra"].lower()
    
    radio = dato["Diametro_cm"] / 2
    area_cm2 = pi * (radio ** 2)
    
    # Agregamos la Biomasa al diccionario para poder usarla al exportar
    Geometria_Pellets[(tipo, muestra)] = {
        "Biomasa": dato["Biomasa"],
        "A": area_cm2, 
        "D": dato["Distancia_cm"]
    }

print("Geometrías cargadas exitosamente en memoria.\n")

# ACUMULADOR GLOBAL: Aquí guardaremos las tablas de todos los archivos

for config in Configuracion_Biomasas:
    print("=" * 80)
    print(f"INICIANDO ANÁLISIS DE LA BIOMASA: {config['Nombre']}")
    print("=" * 80)
    
    # Verificación de entorno: Si la carpeta de imágenes no existe, Python la crea
    if not os.path.exists(config["Ruta_Imagenes"]):
        os.makedirs(config["Ruta_Imagenes"])
        print(f"Creada carpeta de salida visual: {config['Ruta_Imagenes']}")
    
    # --- NUEVO: Verificación y creación de la carpeta "Diagnostico" ---
    Ruta_Diagnostico = os.path.join(config["Ruta_Base"], "Diagnostico")
    if not os.path.exists(Ruta_Diagnostico):
        os.makedirs(Ruta_Diagnostico)
        print(f"Creada carpeta de diagnóstico: {Ruta_Diagnostico}")

    # Buscamos los archivos específicos de la biomasa actual
    CarpetaMediciones = os.path.join(config["Ruta_Base"], "*.xls")
    Archivos = glob.glob(CarpetaMediciones)
    print(f"Archivos válidos encontrados para procesar: {len(Archivos)}\n")

    Reporte_Global = []



    for i in Archivos: 

        ArchivosLower = i.lower() 
        Informacion = os.path.basename(ArchivosLower) 

        NombreLimpio = Informacion.replace(".xls", "").replace("cf-", "")
        Nombres = NombreLimpio.split("_")

        tipo_archivo = Nombres[0].strip() 
        numero_str = Nombres[1].strip()   
        
        if "1" in numero_str:
            muestra_archivo = "1er"
        elif "2" in numero_str:
            muestra_archivo = "2do"
        else:
            muestra_archivo = numero_str 
            
        old_stderr = sys.stderr
        old_stdout = sys.stdout
        sys.stdout = io.StringIO() 
        sys.stderr = io.StringIO()
        
        ExcelEnMemoria = pd.ExcelFile(i) 

        sys.stderr = old_stderr
        sys.stdout = old_stdout

        pestagnas = ["Data"] + [f"Append{j}" for j in range(1, 10)] 

        Bloque_Hojas_Pellet = []
        
        for pestagna in pestagnas: 
            try: 
                old_stderr = sys.stderr
                sys.stderr = io.StringIO()

                df_Hoja = pd.read_excel(ExcelEnMemoria, sheet_name=pestagna) 

                sys.stderr = old_stderr

                ValoresFiltrados = df_Hoja[df_Hoja["F_AB"] >= 8000].copy() 

                Columnas_Base = ValoresFiltrados[["F_AB", "Gp_AB", "Cp_AB"]].copy() 
                Columnas_Base["Gp_AB"] = Columnas_Base["Gp_AB"].abs() 
                Columnas_Base["Cp_AB"] = Columnas_Base["Cp_AB"].abs() 

                Bloque_Hojas_Pellet.append(Columnas_Base) 

            except Exception as e:
                sys.stderr = old_stderr
                continue


    # PROMEDIADO Y CÁLCULO DE ERRORES
        if Bloque_Hojas_Pellet:
            Tabla_Temporal = pd.concat(Bloque_Hojas_Pellet, ignore_index=True)
            
            # 1. Agrupamos por frecuencia y sacamos media y desviación estándar simultáneamente
            Muestra_Agrupada = Tabla_Temporal.groupby("F_AB").agg(
                MeanGp=("Gp_AB", "mean"),
                StdGp=("Gp_AB", "std"),
                MeanCp=("Cp_AB", "mean"),
                StdCp=("Cp_AB", "std")
            ).reset_index()

            Freq = Muestra_Agrupada["F_AB"]
            MeanGp = Muestra_Agrupada["MeanGp"]
            MeanCp = Muestra_Agrupada["MeanCp"]
            StdGp = Muestra_Agrupada["StdGp"]
            StdCp = Muestra_Agrupada["StdCp"]

            # --- EXTRACCIÓN DINÁMICA DE LA GEOMETRÍA ---
            if (tipo_archivo, muestra_archivo) in Geometria_Pellets:
                Biomasa = Geometria_Pellets[(tipo_archivo, muestra_archivo)]["Biomasa"]
                A = Geometria_Pellets[(tipo_archivo, muestra_archivo)]["A"]
                D = Geometria_Pellets[(tipo_archivo, muestra_archivo)]["D"]
            else:
                print(f"ALERTA: El pellet '{tipo_archivo}_{muestra_archivo}' no está en la tabla. Saltando...")
                continue 

    # --- ECUACIONES FÍSICAS EXACTAS (Valores Centrales) ---
            TangentLoss = MeanGp / (2 * pi * Freq * MeanCp)
            RealE = (MeanCp * D) / (EPSILON_0 * A)
            ImaE = (MeanGp * D) / (2 * pi * Freq * A * EPSILON_0)
            
            # NUEVO: Cálculo de Conductividad AC (S/cm)
            Conductividad = MeanGp * (D / A)
            # Descomenta la siguiente línea si necesitas obligatoriamente S/mm
            # Conductividad = Conductividad * 0.1 

            # --- CÁLCULO DE DESVIACIONES ESTÁNDAR Y ERROR ESTÁNDAR (SEM) ---
            N = 10 # Tus 10 réplicas (Data + 9 Appends)
            Raiz_N = np.sqrt(N)

            # Errores para Epsilon Real
            Std_RealE = (StdCp * D) / (EPSILON_0 * A)
            Error_RealE = Std_RealE / Raiz_N

            # Errores para Epsilon Imaginario
            Std_ImaE = (StdGp * D) / (2 * pi * Freq * A * EPSILON_0)
            Error_ImaE = Std_ImaE / Raiz_N

            # NUEVO: Errores para la Conductividad (Dependencia lineal de Gp)
            Std_Conductividad = StdGp * (D / A)
            # Si activaste S/mm arriba, debes hacer: Std_Conductividad = Std_Conductividad * 0.1
            Error_Conductividad = Std_Conductividad / Raiz_N

            # Errores para Tangente de Pérdidas
            Error_Relativo_Gp = StdGp / MeanGp
            Error_Relativo_Cp = StdCp / MeanCp
            Std_TanDelta = TangentLoss * np.sqrt(Error_Relativo_Gp**2 + Error_Relativo_Cp**2)
            Error_TanDelta = Std_TanDelta / Raiz_N

            # Aqui inicia el cambio en el codigo 
            # Esta parte del codigo debe quedar igual 

    # --- ARMADO DEL DATAFRAME HORIZONTAL ---
            sufijo = f"{tipo_archivo.upper()}_{muestra_archivo}"

            df_pellet_resultado = pd.DataFrame({
                f'Biomasa_{sufijo}': [Biomasa] * len(Freq),
                f'Tipo_{sufijo}': [tipo_archivo.upper()] * len(Freq),
                f'Muestra_{sufijo}': [muestra_archivo] * len(Freq),
                f'Frecuencia_Hz_{sufijo}': Freq.values,
                
                # --- NUEVO: DATOS CRUDOS SIN CONVERTIR PARA TROUBLESHOOTING ---
                f'Cp_Raw_{sufijo}': MeanCp.values,
                f'Gp_Raw_{sufijo}': MeanGp.values,
                
                # Valores Centrales calculados
                f'Tan_Delta_{sufijo}': TangentLoss.values,
                f'Epsilon_Real_{sufijo}': RealE.values,
                f'Epsilon_Imag_{sufijo}': ImaE.values,
                f'Conductividad_{sufijo}': Conductividad.values,
                
                # (El resto de tus columnas de errores y desviaciones estándar se quedan exactamente igual...)

                # Desviación Estándar 
                f'Std_TanDelta_{sufijo}': Std_TanDelta.values,
                f'Std_EpsilonReal_{sufijo}': Std_RealE.values,
                f'Std_EpsilonImag_{sufijo}': Std_ImaE.values,
                f'Std_Conductividad_{sufijo}': Std_Conductividad.values, # NUEVO
                
                # Error Estándar 
                f'Error_TanDelta_{sufijo}': Error_TanDelta.values,
                f'Error_EpsilonReal_{sufijo}': Error_RealE.values,
                f'Error_EpsilonImag_{sufijo}': Error_ImaE.values,
                f'Error_Conductividad_{sufijo}': Error_Conductividad.values # NUEVO
            })
            
            Reporte_Global.append(df_pellet_resultado)
            print(f"-> Datos y Errores de {tipo_archivo.upper()} ({muestra_archivo}) listos para consolidación.")


# ==============================================================================
# CONSOLIDACIÓN Y EXPORTACIÓN FINAL A EXCEL (Fuera de los bucles)
# ==============================================================================
    if Reporte_Global:
        
        print(f"\nConcatenando pellets de {config['Nombre']} horizontalmente...")
        Mega_Tabla_Final = pd.concat(Reporte_Global, axis=1)
            
        # Ruta dinámica del Excel basada en la carpeta activa
        Ruta_Salida = os.path.join(config["Ruta_Base"], config["Nombre_Excel"])
        Mega_Tabla_Final.to_excel(Ruta_Salida, index=False)
        print(f"¡Reporte Excel consolidado guardado en:\n{Ruta_Salida}")
    else:
        print("No se encontraron datos procesados para exportar.")

    # ==============================================================================
    # MÓDULO DE DIAGNÓSTICO: GUARDADO AUTOMÁTICO
    # ==============================================================================
    print(f"\nGenerando y guardando gráficas de diagnóstico instrumental para {config['Nombre']}...")

    sufijos_diagnostico = [col.replace('Cp_Raw_', '') for col in Mega_Tabla_Final.columns if col.startswith('Cp_Raw_')]

    for sufijo in sufijos_diagnostico:
        freq = Mega_Tabla_Final[f"Frecuencia_Hz_{sufijo}"]
        cp_cruda = Mega_Tabla_Final[f"Cp_Raw_{sufijo}"]
        gp_cruda = Mega_Tabla_Final[f"Gp_Raw_{sufijo}"]
        etiqueta = sufijo.replace("_", " ")
        
        # --- Gráfica A: Capacitancia ---
        fig_cp = plt.figure(figsize=(7, 4))
        plt.loglog(freq, cp_cruda, '-o', color='tab:blue', markersize=4, label=f"Cp cruda ({etiqueta})")
        plt.xlabel('Frequency [Hz]', fontsize=11)
        plt.ylabel('Raw Capacitance Cp [F]', fontsize=11)
        plt.title(f'Control de Calidad: Cp (Log-Log) - {etiqueta}')
        plt.grid(True, which="both", ls="--", alpha=0.5)
        plt.legend()
        plt.tight_layout()
        
        # Guardado automático y cierre
        nombre_archivo_cp = os.path.join(Ruta_Diagnostico, f"Diagnostico_Cp_{sufijo}.jpg")
        plt.savefig(nombre_archivo_cp, dpi=200, bbox_inches='tight')
        plt.close(fig_cp)
        
        # --- Gráfica B: Conductancia ---
        fig_gp = plt.figure(figsize=(7, 4))
        plt.semilogx(freq, gp_cruda, '-s', color='tab:orange', markersize=4, label=f"Gp cruda ({etiqueta})")
        plt.xlabel('Frequency [Hz]', fontsize=11)
        plt.ylabel('Raw Conductance Gp [S]', fontsize=11)
        plt.title(f'Control de Calidad: Gp (Semi-Log) - {etiqueta}')
        plt.grid(True, which="both", ls="--", alpha=0.5)
        plt.legend()
        plt.tight_layout()
        
        # Guardado automático y cierre
        nombre_archivo_gp = os.path.join(Ruta_Diagnostico, f"Diagnostico_Gp_{sufijo}.jpg")
        plt.savefig(nombre_archivo_gp, dpi=200, bbox_inches='tight')
        plt.close(fig_gp)

        # ==============================================================================
        # MÓDULO DE DIAGNÓSTICO: GUARDADO AUTOMÁTICO
        # ==============================================================================
        print(f"\nGenerando y guardando gráficas de diagnóstico instrumental para {config['Nombre']}...")

        sufijos_diagnostico = [col.replace('Cp_Raw_', '') for col in Mega_Tabla_Final.columns if col.startswith('Cp_Raw_')]

        for sufijo in sufijos_diagnostico:
            freq = Mega_Tabla_Final[f"Frecuencia_Hz_{sufijo}"]
            cp_cruda = Mega_Tabla_Final[f"Cp_Raw_{sufijo}"]
            gp_cruda = Mega_Tabla_Final[f"Gp_Raw_{sufijo}"]
            etiqueta = sufijo.replace("_", " ")
            
            # --- Gráfica A: Capacitancia ---
            fig_cp = plt.figure(figsize=(7, 4))
            plt.loglog(freq, cp_cruda, '-o', color='tab:blue', markersize=4, label=f"Cp cruda ({etiqueta})")
            plt.xlabel('Frequency [Hz]', fontsize=11)
            plt.ylabel('Raw Capacitance Cp [F]', fontsize=11)
            plt.title(f'Control de Calidad: Cp (Log-Log) - {etiqueta}')
            plt.grid(True, which="both", ls="--", alpha=0.5)
            plt.legend()
            plt.tight_layout()
            
            # Guardado automático y cierre
            nombre_archivo_cp = os.path.join(Ruta_Diagnostico, f"Diagnostico_Cp_{sufijo}.jpg")
            plt.savefig(nombre_archivo_cp, dpi=200, bbox_inches='tight')
            plt.close(fig_cp)
            
            # --- Gráfica B: Conductancia ---
            fig_gp = plt.figure(figsize=(7, 4))
            plt.semilogx(freq, gp_cruda, '-s', color='tab:orange', markersize=4, label=f"Gp cruda ({etiqueta})")
            plt.xlabel('Frequency [Hz]', fontsize=11)
            plt.ylabel('Raw Conductance Gp [S]', fontsize=11)
            plt.title(f'Control de Calidad: Gp (Semi-Log) - {etiqueta}')
            plt.grid(True, which="both", ls="--", alpha=0.5)
            plt.legend()
            plt.tight_layout()
            
            # Guardado automático y cierre
            nombre_archivo_gp = os.path.join(Ruta_Diagnostico, f"Diagnostico_Gp_{sufijo}.jpg")
            plt.savefig(nombre_archivo_gp, dpi=200, bbox_inches='tight')
            plt.close(fig_gp)

        # ==============================================================================
        # MÓDULO DE VISUALIZACIÓN CIENTÍFICA (GRÁFICAS AUTOMATIZADAS)
        # ==============================================================================
        print(f"\nIniciando generación de gráficas científicas para {config['Nombre']}...")

        # Preparamos los 3 lienzos de forma independiente antes de dibujar
        fig_eps, ax_eps = plt.subplots(figsize=(8, 6))
        fig_sig, ax_sig = plt.subplots(figsize=(8, 6))
        fig_tan, ax_tan = plt.subplots(figsize=(8, 6))

        # Filtramos los sufijos directamente de la tabla actual
        sufijos_ciencia = [col.replace('Tipo_', '') for col in Mega_Tabla_Final.columns if col.startswith('Tipo_')]

        for sufijo in sufijos_ciencia:
            # Extracción de vectores
            freq = Mega_Tabla_Final[f"Frecuencia_Hz_{sufijo}"]
            
            eps = Mega_Tabla_Final[f"Epsilon_Real_{sufijo}"]
            err_eps = Mega_Tabla_Final[f"Error_EpsilonReal_{sufijo}"]
            
            sig = Mega_Tabla_Final[f"Conductividad_{sufijo}"]
            err_sig = Mega_Tabla_Final[f"Error_Conductividad_{sufijo}"]
            
            tan = Mega_Tabla_Final[f"Tan_Delta_{sufijo}"]
            err_tan = Mega_Tabla_Final[f"Error_TanDelta_{sufijo}"]
            
            etiqueta = sufijo.replace("_", " ")
            
            # Dibujamos las curvas (Esto se repite por cada pellet)
            ax_eps.errorbar(freq, eps, yerr=err_eps, fmt='-o', markersize=4, capsize=3, alpha=0.8, label=etiqueta)
            ax_sig.errorbar(freq, sig, yerr=err_sig, fmt='-s', markersize=4, capsize=3, alpha=0.8, label=etiqueta)
            ax_tan.errorbar(freq, tan, yerr=err_tan, fmt='-^', markersize=4, capsize=3, alpha=0.8, label=etiqueta)
                    
        # --- Formato Estético (FUERA del bucle de sufijos, se ejecuta una sola vez por biomasa) ---
        biomasa = config["Nombre"]

        ax_eps.set_xscale('log')
        ax_eps.set_xlabel('Frequency [Hz]', fontsize=12)
        ax_eps.set_ylabel(r'$\epsilon_r$', fontsize=14)
        ax_eps.set_title(f'Permitividad Relativa - {biomasa}')
        ax_eps.legend(fontsize=9)
        ax_eps.grid(True, which="both", ls="--", alpha=0.4)

        ax_sig.set_xscale('log')
        ax_sig.set_xlabel('Frequency [Hz]', fontsize=12)
        ax_sig.set_ylabel(r'$\sigma_{ac}$ [S/cm]', fontsize=14)
        ax_sig.set_title(f'Conductividad AC - {biomasa}')
        ax_sig.legend(fontsize=9)
        ax_sig.grid(True, which="both", ls="--", alpha=0.4)

        ax_tan.set_xscale('log')
        ax_tan.set_xlabel('Frequency [Hz]', fontsize=12)
        ax_tan.set_ylabel(r'$\tan(\delta)$', fontsize=14)
        ax_tan.set_title(f'Tangente de Pérdidas - {biomasa}')
        ax_tan.legend(fontsize=9)
        ax_tan.grid(True, which="both", ls="--", alpha=0.4)

        # --- Guardado Automático Dinámico (FUERA del bucle de sufijos) ---
        nombre_archivo = biomasa.replace(" ", "_")
        ruta_guardado = config["Ruta_Imagenes"]

        fig_eps.savefig(os.path.join(ruta_guardado, f"Permittivity_{nombre_archivo}.jpg"), dpi=300, bbox_inches='tight')
        fig_sig.savefig(os.path.join(ruta_guardado, f"Conductivity_{nombre_archivo}.jpg"), dpi=300, bbox_inches='tight')
        fig_tan.savefig(os.path.join(ruta_guardado, f"LossTangent_{nombre_archivo}.jpg"), dpi=300, bbox_inches='tight')

        # Liberamos memoria
        plt.close(fig_eps)
        plt.close(fig_sig)
        plt.close(fig_tan)