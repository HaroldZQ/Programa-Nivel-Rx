import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class LinkBudgetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("APU SPACE // Mission Control")
        self.root.geometry("640x720")
        self.root.resizable(False, False)
        self.root.configure(bg="#07111F")

        # Style
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TLabel", background="#0B1726", foreground="#D8E7F2")
        self.style.configure("TLabelFrame", background="#0B1726", foreground="#F7B955")
        self.style.configure("TFrame", background="#0B1726")
        self.style.configure(
            "Mission.TEntry",
            fieldbackground="#252B33",
            foreground="#F2F4F7",
            bordercolor="#B8893D",
            lightcolor="#B8893D",
            darkcolor="#151A20",
        )

        # Variables
        self.csv_path = tk.StringVar()
        
        # Parámetros RF por defecto (tomados de la imagen)
        self.freq_mhz = tk.DoubleVar(value=433.0)
        self.offset_gs_m = tk.DoubleVar(value=150.0)
        self.ptx_dbm = tk.DoubleVar(value=30.0)
        self.gtx_dbi = tk.DoubleVar(value=2.5)
        self.grx_dbi = tk.DoubleVar(value=2.5)
        self.lmisc_db = tk.DoubleVar(value=2.0)
        self.srx_2_4k_dbm = tk.DoubleVar(value=-147.0)
        self.srx_0_3k_dbm = tk.DoubleVar(value=-145.0)

        self.create_widgets()

    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#07111F", pady=14)
        header_frame.pack(fill=tk.X)
        lbl_title = tk.Label(header_frame, text="APU SPACE // MISSION CONTROL", 
                     font=("Consolas", 15, "bold"), fg="#F7B955", bg="#07111F")
        lbl_title.pack()
        lbl_sub = tk.Label(header_frame, text="FLIGHT TELEMETRY  /  LoRa 433 MHz  /  OPENROCKET", 
                   font=("Consolas", 9), fg="#6EDFF6", bg="#07111F")
        lbl_sub.pack()

        # 1. Carga de Archivo
        frame_file = ttk.LabelFrame(self.root, text=" 1. Archivo de Simulación de Vuelo (OpenRocket) ", padding=10)
        frame_file.pack(fill=tk.X, padx=15, pady=10)

        entry_file = ttk.Entry(frame_file, textvariable=self.csv_path, width=50, style="Mission.TEntry")
        entry_file.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        btn_browse = ttk.Button(frame_file, text="Examinar CSV", command=self.browse_file)
        btn_browse.pack(side=tk.RIGHT, padx=5)

        # 2. Parámetros del Sistema de Radiofrecuencia (Añadidos por usuario)
        frame_params = ttk.LabelFrame(self.root, text=" 2. Parámetros del Sistema de RF (Modificables) ", padding=10)
        frame_params.pack(fill=tk.X, padx=15, pady=5)

        params_def = [
            ("Frecuencia de Operación (f):", self.freq_mhz, "MHz"),
            ("Distancia Estación Terrena a Rampa:", self.offset_gs_m, "m"),
            ("Potencia de Transmisión (Ptx):", self.ptx_dbm, "dBm"),
            ("Ganancia Antena TX (GTx):", self.gtx_dbi, "dBi"),
            ("Ganancia Antena RX (GRx):", self.grx_dbi, "dBi"),
            ("Pérdidas Misceláneas (Lmisc):", self.lmisc_db, "dB"),
            ("Sensibilidad RX @ 2.4 kbps (Srx):", self.srx_2_4k_dbm, "dBm"),
            ("Sensibilidad RX @ 0.3 kbps (Srx):", self.srx_0_3k_dbm, "dBm"),
        ]

        for row_idx, (label_text, var, unit) in enumerate(params_def):
            lbl = ttk.Label(frame_params, text=label_text)
            lbl.grid(row=row_idx, column=0, sticky=tk.W, pady=3, padx=5)
            entry = ttk.Entry(frame_params, textvariable=var, width=12, justify="right", style="Mission.TEntry")
            entry.grid(row=row_idx, column=1, sticky=tk.E, pady=3, padx=5)
            lbl_u = ttk.Label(frame_params, text=unit)
            lbl_u.grid(row=row_idx, column=2, sticky=tk.W, pady=3, padx=2)

        # 3. Botón de Acción
        btn_run = tk.Button(self.root, text="▶  EJECUTAR ANÁLISIS DE MISIÓN", 
                    font=("Consolas", 11, "bold"), bg="#D88722", fg="#07111F", 
                    activebackground="#F7B955", activeforeground="#07111F",
                            cursor="hand2", pady=8, command=self.process_data)
        btn_run.pack(fill=tk.X, padx=15, pady=12)

        # Consola de estado / log
        self.txt_log = tk.Text(self.root, height=10, bg="#06101C", fg="#9BE7F5", insertbackground="#F7B955", font=("Consolas", 9), relief=tk.FLAT)
        self.txt_log.pack(fill=tk.BOTH, padx=15, pady=5, expand=True)
        self.log("Programa iniciado. Selecciona archivo CSV de OpenRocket para comenzar.")

    def log(self, text):
        self.txt_log.insert(tk.END, text + "\n")
        self.txt_log.see(tk.END)

    def browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Seleccionar archivo CSV de OpenRocket",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        if file_selected:
            self.csv_path.set(file_selected)
            self.log(f"Archivo cargado: {os.path.basename(file_selected)}")

    def process_data(self):
        csv_file = self.csv_path.get()
        if not csv_file or not os.path.exists(csv_file):
            messagebox.showerror("Error", "Por favor selecciona un archivo CSV válido de OpenRocket.")
            return

        try:
            self.log("\n--> Leyendo telemetría...")
            headers = [
                "Tiempo (s)", "Altitud (m)", "Velocidad vertical (m/s)", "Aceleración vertical (m/s²)",
                "Velocidad total (m/s)", "Aceleración total (m/s²)", "Posición Este (m)", "Posición Norte (m)",
                "Distancia lateral (m)", "Velocidad horizontal (m/s)", "Aceleración horizontal (m/s²)",
                "Latitud (°)", "Longitud (°)"
            ]
            df = pd.read_csv(csv_file, comment='#', header=None)
            df.columns = headers

            f = self.freq_mhz.get()
            d_offset = self.offset_gs_m.get()
            ptx = self.ptx_dbm.get()
            gtx = self.gtx_dbi.get()
            grx = self.grx_dbi.get()
            lmisc = self.lmisc_db.get()
            srx_2_4 = self.srx_2_4k_dbm.get()
            srx_0_3 = self.srx_0_3k_dbm.get()

            # Cálculos
            self.log("--> Calculando parámetros de enlace geométrico y RF...")
            df["Dist. Cohete a GS (m)"] = np.sqrt(df["Altitud (m)"]**2 + (d_offset + df["Distancia lateral (m)"])**2)
            df["Dist. Cohete a GS (km)"] = df["Dist. Cohete a GS (m)"] / 1000.0

            cte_fspl = 20.0 * np.log10(f) + 32.44
            df["FSPL (dBm)"] = 20.0 * np.log10(df["Dist. Cohete a GS (km)"]) + cte_fspl
            df["Prx (dBm)"] = (ptx + gtx + grx - lmisc) - df["FSPL (dBm)"]

            df["Margen 2.4k (dB)"] = df["Prx (dBm)"] - srx_2_4
            df["Margen 0.3k (dB)"] = df["Prx (dBm)"] - srx_0_3

            idx_ap = df["Altitud (m)"].idxmax()
            t_ap = df.loc[idx_ap, "Tiempo (s)"]
            alt_ap = df.loc[idx_ap, "Altitud (m)"]
            dist_ap = df.loc[idx_ap, "Dist. Cohete a GS (km)"]
            prx_ap = df.loc[idx_ap, "Prx (dBm)"]

            self.log(f"  * Apogeo máximo: {alt_ap:.2f} m a los {t_ap:.2f} s")
            self.log(f"  * Prx en Apogeo: {prx_ap:.2f} dBm (Margen @ 2.4k: {df.loc[idx_ap, 'Margen 2.4k (dB)']:.2f} dB)")
            self.log(f"  * Prx Mínima en todo el vuelo: {df['Prx (dBm)'].min():.2f} dBm")

            # Generar Excel
            output_dir = os.path.dirname(csv_file)
            base_name = os.path.splitext(os.path.basename(csv_file))[0]
            xlsx_out = os.path.join(output_dir, f"{base_name}_LinkBudget_Calculado.xlsx")
            png_out = os.path.join(output_dir, f"{base_name}_Dashboard_Graficas.png")
            html_out = os.path.join(output_dir, f"{base_name}_Dashboard_Graficas_Interactivo.html")

            self.log(f"--> Exportando archivo Excel a: {os.path.basename(xlsx_out)}...")
            self.export_excel(df, xlsx_out, f, d_offset, ptx, gtx, grx, lmisc, srx_2_4, srx_0_3, alt_ap, t_ap, prx_ap)

            # Generar Gráficas
            self.log(f"--> Generando imagen de dashboard a: {os.path.basename(png_out)}...")
            self.plot_dashboard(df, png_out, t_ap, prx_ap, dist_ap)

            self.log(f"--> Generando dashboard interactivo a: {os.path.basename(html_out)}...")
            self.plot_interactive_dashboard(df, html_out, t_ap, prx_ap, dist_ap)

            self.log("✓ ¡PROCESO COMPLETADO EXITOSAMENTE!")
            messagebox.showinfo("Éxito", f"Cálculo completado.\n\nArchivos generados:\n- {os.path.basename(xlsx_out)}\n- {os.path.basename(png_out)}\n- {os.path.basename(html_out)}\n\nEl dashboard interactivo se abrió en el navegador.")

        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            messagebox.showerror("Error durante el cálculo", str(e))

    def export_excel(self, df, filename, f, d_offset, ptx, gtx, grx, lmisc, srx_2_4, srx_0_3, alt_ap, t_ap, prx_ap):
        cols_modelo = [
            "Tiempo (s)", "Altitud (m)", "Distancia lateral (m)", "Dist. Cohete a GS (m)",
            "Dist. Cohete a GS (km)", "Velocidad total (m/s)", "Latitud (°)", "Longitud (°)",
            "FSPL (dBm)", "Prx (dBm)"
        ]
        df_export = df[cols_modelo].copy()
        df_export.rename(columns={
            "Distancia lateral (m)": "Dist. Lateral Cohete (m)",
            "Velocidad total (m/s)": "Velocidad (m/s)"
        }, inplace=True)

        wb = openpyxl.Workbook()
        ws_res = wb.active
        ws_res.title = f"Resumen Link Budget (GS {int(d_offset)}m)"
        ws_tel = wb.create_sheet(title="Telemetria & Link Budget")

        # Parámetros RF en Resumen
        ws_res["A1"] = f"ANÁLISIS DE ENLACE DE TELEMETRÍA LoRa {int(f)} MHz (GS A {int(d_offset)}m)"
        ws_res["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
        ws_res["A2"] = f"Estación Terrena ubicada a {int(d_offset)} m de la rampa de lanzamiento (Zona Segura)"
        ws_res["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

        ws_res["A4"] = "1. PARÁMETROS DEL SISTEMA DE RADIOFRECUENCIA (RF)"
        ws_res["A4"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        ws_res["A4"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

        params = [
            ("Frecuencia de Operación (f)", f, "MHz", "Banda ISM UHF"),
            ("Distancia Estación Terrena a Rampa", d_offset, "m", "Radio de seguridad perimetral"),
            ("Potencia de Transmisión (Ptx)", ptx, "dBm (1000 mW)", "Módulo E32-433T30D (SX1278)"),
            ("Ganancia Antena TX (Gtx) [Tx433-jk-11]", gtx, "dBi", "Antena omnidireccional TX433-JK-11"),
            ("Ganancia Antena RX (Grx) [Tx433-jk-11]", grx, "dBi", "Antena omnidireccional TX433-JK-11"),
            ("Pérdidas Misceláneas (Lmisc)", lmisc, "dB", "Pérdida por conectores SMA, coaxial y orientación"),
            ("PIRE / EIRP (Ptx + Gtx)", ptx + gtx, "dBm", "Potencia Isotrópica Radiada Equivalente"),
            ("Receiving Sensitivity @ 2.4 kbps (Srx)-E90-DTU(433L30)-V8 (Tecnologia LORA)", srx_2_4, "dBm", "Air Data Rate estándar / recomendado"),
            ("Receiving Sensitivity @ 0.3 kbps (Srx)-E90-DTU(433L30)-V8 (Tecnologia LORA)", srx_0_3, "dBm", "Air Data Rate modo ultra largo alcance")
        ]

        for i, (k, v, u, desc) in enumerate(params, start=5):
            ws_res.cell(row=i, column=1, value=k).font = Font(name="Calibri", bold=True)
            ws_res.cell(row=i, column=2, value=v)
            ws_res.cell(row=i, column=3, value=u)
            ws_res.cell(row=i, column=4, value=desc)

        # Telemetria & Link Budget
        ws_tel.append(["NOTA: Datos obtenidos de simulación OpenRocket y procesados automáticamente"])
        ws_tel.append(list(df_export.columns))
        for row in df_export.itertuples(index=False):
            ws_tel.append(list(row))

        wb.save(filename)

    def plot_dashboard(self, df, filename, t_ap, prx_ap, dist_ap):
        plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")
        fig = plt.figure(figsize=(15, 9.5), facecolor="#F8FAFC")

        # 1. Prx vs Tiempo
        ax1 = plt.subplot(2, 2, 1)
        ax1.plot(df["Tiempo (s)"], df["Prx (dBm)"], color="#D9381E", linewidth=2.4, label="$P_{rx}$ [dBm]")
        ax1.scatter(t_ap, prx_ap, color="#1F4E78", s=85, zorder=5, 
                    label=f"Apogeo (t={t_ap:.1f} s, {prx_ap:.2f} dBm)")
        ax1.set_title("Potencia Recibida ($P_{rx}$) vs. Tiempo de Vuelo", fontsize=12, fontweight="bold", color="#1E293B", pad=10)
        ax1.set_xlabel("Tiempo de Vuelo (s)", fontsize=10, fontweight="bold")
        ax1.set_ylabel("Potencia Recibida $P_{rx}$ (dBm)", fontsize=10, fontweight="bold")
        ax1.legend(loc="lower right", frameon=True, facecolor="white")
        ax1.grid(True, linestyle="--", alpha=0.55)

        # 2. Prx vs Distancia 3D
        ax2 = plt.subplot(2, 2, 2)
        ax2.plot(df["Dist. Cohete a GS (km)"], df["Prx (dBm)"], color="#1F4E78", linewidth=2.4)
        ax2.scatter(df["Dist. Cohete a GS (km)"].iloc[0], df["Prx (dBm)"].iloc[0], color="#10B981", s=85, zorder=5,
                    label=f"Rampa (d={df['Dist. Cohete a GS (km)'].iloc[0]:.3f} km)")
        ax2.scatter(dist_ap, prx_ap, color="#D9381E", s=85, zorder=5,
                    label=f"Apogeo (d={dist_ap:.3f} km)")
        ax2.set_title("Potencia Recibida ($P_{rx}$) vs. Distancia 3D a GS (km)", fontsize=12, fontweight="bold", color="#1E293B", pad=10)
        ax2.set_xlabel("Distancia Tridimensional a GS (km)", fontsize=10, fontweight="bold")
        ax2.set_ylabel("Potencia Recibida $P_{rx}$ (dBm)", fontsize=10, fontweight="bold")
        ax2.legend(loc="lower left", frameon=True, facecolor="white")
        ax2.grid(True, linestyle="--", alpha=0.55)

        # 3. Altitud vs Margen
        ax3 = plt.subplot(2, 2, 3)
        ax3_twin = ax3.twinx()
        l1 = ax3.plot(df["Tiempo (s)"], df["Altitud (m)"], color="#0284C7", linewidth=2.2, label="Altitud (m)")
        l2 = ax3_twin.plot(df["Tiempo (s)"], df["Margen 2.4k (dB)"], color="#9333EA", linewidth=2.0, linestyle="--", label="Margen LoRa @ 2.4k (dB)")
        ax3.set_title("Perfil de Altitud vs. Margen de Enlace", fontsize=12, fontweight="bold", color="#1E293B", pad=10)
        ax3.set_xlabel("Tiempo de Vuelo (s)", fontsize=10, fontweight="bold")
        ax3.set_ylabel("Altitud (m)", fontsize=10, color="#0284C7", fontweight="bold")
        ax3_twin.set_ylabel("Margen de Enlace (dB)", fontsize=10, color="#9333EA", fontweight="bold")
        lines = l1 + l2
        ax3.legend(lines, [l.get_label() for l in lines], loc="center right", frameon=True, facecolor="white")
        ax3.grid(True, linestyle="--", alpha=0.55)

        # 4. Trayectoria 2D
        ax4 = plt.subplot(2, 2, 4)
        sc = ax4.scatter(df["Distancia lateral (m)"], df["Altitud (m)"], c=df["Prx (dBm)"], cmap="plasma", s=22, alpha=0.85)
        cbar = plt.colorbar(sc, ax=ax4, pad=0.02)
        cbar.set_label("Nivel de Señal $P_{rx}$ (dBm)", fontsize=10, fontweight="bold")
        ax4.set_title("Trayectoria 2D con Atenuación de Señal LoRa", fontsize=12, fontweight="bold", color="#1E293B", pad=10)
        ax4.set_xlabel("Deriva Lateral desde la Rampa (m)", fontsize=10, fontweight="bold")
        ax4.set_ylabel("Altitud (m)", fontsize=10, fontweight="bold")
        ax4.grid(True, linestyle="--", alpha=0.55)

        plt.suptitle("TELEMETRÍA Y LINK BUDGET ROCKET APU SPACE", fontsize=14, fontweight="bold", color="#0F172A", y=0.98)
        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(filename, dpi=300, facecolor=fig.get_facecolor())
        plt.close()

    def plot_interactive_dashboard(self, df, filename, t_ap, prx_ap, dist_ap):
        fig = make_subplots(
            rows=2,
            cols=2,
            specs=[[{}, {}], [{"secondary_y": True}, {}]],
            subplot_titles=(
                "Potencia Recibida vs. Tiempo",
                "Potencia Recibida vs. Distancia 3D",
                "Perfil de Altitud y Margen de Enlace",
                "Trayectoria 2D con Atenuación",
            ),
            vertical_spacing=0.12,
            horizontal_spacing=0.18,
        )

        time = df["Tiempo (s)"]
        prx = df["Prx (dBm)"]
        distance = df["Dist. Cohete a GS (km)"]
        hover_text = (
            "Tiempo: " + time.round(3).astype(str) + " s<br>"
            + "Altitud: " + df["Altitud (m)"].round(2).astype(str) + " m<br>"
            + "Distancia a GS: " + distance.round(4).astype(str) + " km<br>"
            + "Prx: " + prx.round(2).astype(str) + " dBm<br>"
            + "Margen 2.4k: " + df["Margen 2.4k (dB)"].round(2).astype(str) + " dB"
        )

        fig.add_trace(go.Scatter(
            x=time, y=prx, mode="lines", name="Prx (dBm)",
            line={"color": "#D9381E", "width": 2}, text=hover_text,
            hovertemplate="%{text}<extra></extra>",
        ), row=1, col=1)
        fig.add_trace(go.Scatter(
            x=[t_ap], y=[prx_ap], mode="markers", name="Apogeo",
            marker={"color": "#1F4E78", "size": 10},
            hovertemplate=f"Apogeo<br>Tiempo: {t_ap:.2f} s<br>Prx: {prx_ap:.2f} dBm<extra></extra>",
        ), row=1, col=1)

        fig.add_trace(go.Scatter(
            x=distance, y=prx, mode="lines", name="Prx vs distancia",
            line={"color": "#1F4E78", "width": 2}, text=hover_text,
            hovertemplate="%{text}<extra></extra>", showlegend=False,
        ), row=1, col=2)
        fig.add_trace(go.Scatter(
            x=[distance.iloc[0], dist_ap], y=[prx.iloc[0], prx_ap],
            mode="markers+text", name="Puntos de referencia", text=["Rampa", "Apogeo"],
            textposition="top center", marker={"color": ["#10B981", "#D9381E"], "size": 10},
            hovertemplate="%{text}<br>Distancia: %{x:.3f} km<br>Prx: %{y:.2f} dBm<extra></extra>",
        ), row=1, col=2)

        fig.add_trace(go.Scatter(
            x=time, y=df["Altitud (m)"], mode="lines", name="Altitud (m)",
            line={"color": "#0284C7", "width": 2}, text=hover_text,
            hovertemplate="%{text}<extra></extra>",
        ), row=2, col=1, secondary_y=False)
        fig.add_trace(go.Scatter(
            x=time, y=df["Margen 2.4k (dB)"], mode="lines", name="Margen 2.4k (dB)",
            line={"color": "#9333EA", "width": 2, "dash": "dash"}, text=hover_text,
            hovertemplate="%{text}<extra></extra>",
        ), row=2, col=1, secondary_y=True)

        fig.add_trace(go.Scatter(
            x=df["Distancia lateral (m)"], y=df["Altitud (m)"], mode="markers",
            name="Trayectoria", text=hover_text,
            marker={
                "size": 7, "color": prx, "colorscale": "Plasma",
                "showscale": True, "colorbar": {"title": "Prx (dBm)", "x": 1.02},
            }, hovertemplate="%{text}<extra></extra>", showlegend=False,
        ), row=2, col=2)

        fig.update_xaxes(title_text="Tiempo de vuelo (s)", row=1, col=1)
        fig.update_yaxes(title_text="Prx (dBm)", row=1, col=1)
        fig.update_xaxes(title_text="Distancia 3D Rocket-GS (km)", row=1, col=2)
        fig.update_yaxes(title_text="Prx (dBm)", row=1, col=2)
        fig.update_xaxes(title_text="Tiempo de vuelo (s)", row=2, col=1)
        fig.update_yaxes(
            title_text="Altitud (m)",
            title_standoff=28,
            ticklabelposition="outside",
            automargin=True,
            row=2,
            col=1,
            secondary_y=False,
        )
        fig.update_yaxes(
            title_text="Margen de enlace (dB)",
            title_standoff=32,
            ticklabelposition="outside",
            automargin=True,
            row=2,
            col=1,
            secondary_y=True,
        )
        fig.update_layout(
            yaxis4={
                "anchor": "free",
                "overlaying": "y3",
                "side": "right",
                "position": 0.405,
            }
        )
        fig.update_xaxes(title_text="Deriva lateral (m)", row=2, col=2)
        fig.update_yaxes(title_text="Altitud (m)", row=2, col=2)
        fig.update_layout(
            title="TELEMETRÍA Y LINK BUDGET ROCKET APU SPACE",
            template="plotly_dark", height=850, width=1400,
            paper_bgcolor="#07111F", plot_bgcolor="#0B1A2A",
            font={"family": "Consolas, monospace", "color": "#D8E7F2", "size": 12},
            title_font={"family": "Consolas, monospace", "color": "#F7B955", "size": 20},
            hoverlabel={
                "bgcolor": "#174B5B",
                "bordercolor": "#6EDFF6",
                "font": {"family": "Consolas, monospace", "color": "#F2FAFF", "size": 12},
            },
            hovermode="closest", legend={"orientation": "h", "y": -0.08, "font": {"color": "#D8E7F2"}},
            margin={"l": 115, "r": 145, "t": 90, "b": 80},
        )
        fig.update_xaxes(
            showline=True, linecolor="#31556B", linewidth=1,
            gridcolor="#193448", zerolinecolor="#31556B",
            tickfont={"color": "#A8C4D4"}, title_font={"color": "#6EDFF6"},
        )
        fig.update_yaxes(
            showline=True, linecolor="#31556B", linewidth=1,
            gridcolor="#193448", zerolinecolor="#31556B",
            tickfont={"color": "#A8C4D4"}, title_font={"color": "#6EDFF6"},
        )
        fig.add_annotation(
            text="● LIVE TELEMETRY  //  FLIGHT DATA ANALYSIS",
            x=0, y=1.07, xref="paper", yref="paper", showarrow=False,
            xanchor="left", font={"family": "Consolas, monospace", "size": 11, "color": "#6EDFF6"},
        )
        fig.write_html(
            filename,
            include_plotlyjs=True,
            full_html=True,
            config={
                "displaylogo": False,
                "responsive": True,
                "toImageButtonOptions": {"format": "png", "filename": "link_budget_dashboard", "scale": 2},
            },
            auto_open=False,
        )
        self.add_dashboard_ambient_animation(filename)
        webbrowser.open(os.path.abspath(filename))

    def add_dashboard_ambient_animation(self, filename):
        with open(filename, "r", encoding="utf-8") as dashboard_file:
            html = dashboard_file.read()

        ambient_style = """
<style>
    html, body { background: #07111F !important; }
    body::before {
        content: "";
        position: fixed;
        inset: 0;
        pointer-events: none;
        z-index: 9999;
        opacity: .09;
        background: repeating-linear-gradient(
            0deg, transparent 0, transparent 5px,
            rgba(110, 223, 246, .18) 6px, transparent 7px
        );
        animation: telemetryScan 12s linear infinite;
    }
    body::after {
        content: "APU SPACE  //  FLIGHT DATA LINK  //  SYSTEM NOMINAL";
        position: fixed;
        right: 18px;
        bottom: 10px;
        pointer-events: none;
        color: #6EDFF6;
        font: 10px Consolas, monospace;
        letter-spacing: 1px;
        opacity: .55;
    }
    @keyframes telemetryScan {
        from { transform: translateY(-8px); }
        to { transform: translateY(8px); }
    }
</style>
"""
        html = html.replace("</head>", ambient_style + "</head>", 1)
        with open(filename, "w", encoding="utf-8") as dashboard_file:
            dashboard_file.write(html)

if __name__ == "__main__":
    root = tk.Tk()
    app = LinkBudgetApp(root)
    root.mainloop()
