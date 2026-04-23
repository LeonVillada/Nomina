import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import BOTH, messagebox
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import os
import re

# =========================================================================
# CONFIGURACIÓN GLOBAL
# =========================================================================
DATA_FILE = "employees.xlsx"

# =========================================================================
# UTILIDADES Y VALIDORES
# =========================================================================
class DataValidator:
    """Clase para limpiar y validar datos de entrada."""
    
    @staticmethod
    def clean_numeric(value):
        """
        Limpia un string de caracteres no numéricos, manejando formatos como '2'000.000,00' o '1,500.50'.
        """
        if not value:
            return 0.0
            
        s = str(value).strip()
        # Si tiene Coma Y Punto, asumimos que la coma es decimal si aparece después.
        # Caso común: 1.234.567,89 o 1,234,567.89
        if ',' in s and '.' in s:
            if s.find(',') > s.find('.'):
                # Punto es miles, coma es decimal
                s = s.replace('.', '').replace(',', '.')
            else:
                # Coma es miles, punto es decimal
                s = s.replace(',', '')
        elif ',' in s:
            # Solo coma. Podría ser miles (1,000) o decimal (1,50). 
            # Si hay 3 dígitos después, es probable que sea miles. 
            # Pero para seguridad en nómina, limpiaremos todo lo no numérico excepto el último separador.
            parts = s.split(',')
            if len(parts[-1]) == 3: # 1,000
                s = s.replace(',', '')
            else: # 1,50
                s = s.replace(',', '.')
        
        # Quitar ticks y espacios
        s = s.replace("'", "").replace(" ", "")
        
        # Eliminar cualquier cosa que no sea dígito o el punto que dejamos
        cleaned = re.sub(r'[^0-9.]', '', s)
        
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    @staticmethod
    def is_valid_name(name):
        """Valida que el nombre no sea solo espacios o números."""
        return len(name.strip()) > 2 and not name.isdigit()

# =========================================================================
# MODELOS
# =========================================================================
class Employee:
    """Modelo estructurado de un empleado."""
    def __init__(self, emp_id, nombre, cargo, salario, fecha_ingreso, horario):
        self.emp_id = emp_id
        self.nombre = nombre
        self.cargo = cargo
        self.salario = salario
        self.fecha_ingreso = fecha_ingreso
        self.horario = horario

class PayrollResult:
    """Contenedor para los resultados de un cálculo de nómina."""
    def __init__(self, empleado, total_devengado, salud, pension, total_pagar):
        self.empleado = empleado
        self.total_devengado = total_devengado
        self.salud = salud
        self.pension = pension
        self.total_pagar = total_pagar

# =========================================================================
# SERVICIOS (LÓGICA DE DATOS Y ARCHIVOS)
# =========================================================================
class ExcelService:
    """Servicio para persistencia de datos en Excel."""

    @staticmethod
    def load_employees():
        """Carga los empleados desde el archivo Excel. Crea uno si no existe."""
        if not os.path.exists(DATA_FILE):
            df = pd.DataFrame(columns=[
                "ID", "Nombre", "Cargo",
                "Salario", "FechaIngreso", "Horario"
            ])
            df.to_excel(DATA_FILE, index=False)
            return df
        try:
            # Forzamos ID como string para evitar pérdida de ceros a la izquierda
            return pd.read_excel(DATA_FILE, dtype={"ID": str})
        except Exception as e:
            messagebox.showerror("Error de Archivo", f"No se pudo leer el Excel: {e}")
            return pd.DataFrame()

    @staticmethod
    def save_employees(df):
        """Guarda el DataFrame en el archivo Excel."""
        try:
            df.to_excel(DATA_FILE, index=False)
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Cierra el archivo Excel si está abierto.\n{e}")

class PayrollService:
    """Servicio de cálculo de leyes laborales."""

    @staticmethod
    def calcular_nomina(salario, horas_extra_valor=0):
        """Realiza los cálculos de deducciones legales (Salud 4%, Pensión 4%)."""
        total_devengado = salario + horas_extra_valor
        salud = total_devengado * 0.04
        pension = total_devengado * 0.04
        total_pagar = total_devengado - salud - pension

        return {
            "devengado": total_devengado,
            "salud": salud,
            "pension": pension,
            "neto": total_pagar
        }

class PDFService:
    """Servicio para generación de reportes en PDF."""

    @staticmethod
    def generar_colilla(result: PayrollResult):
        """Crea un documento PDF con el desglose del pago."""
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Encabezado
        pdf.set_fill_color(200, 220, 255)
        pdf.cell(0, 10, "COMPROBANTE DE PAGO - NOMINAMASTER PRO", ln=True, align="C", fill=True)
        pdf.ln(10)
        
        # Datos
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, f"Empleado: {result.empleado}", ln=True)
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 8, f"Fecha de Emision: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True)
        pdf.ln(5)

        # Tabla de conceptos
        pdf.cell(100, 10, "Concepto", 1)
        pdf.cell(0, 10, "Valor", 1, ln=True)
        
        pdf.cell(100, 10, "Salario + Extras", 1)
        pdf.cell(0, 10, f"$ {result.total_devengado:,.0f}", 1, ln=True)
        
        pdf.cell(100, 10, "Deduccion Salud (4%)", 1)
        pdf.cell(0, 10, f"- $ {result.salud:,.0f}", 1, ln=True)
        
        pdf.cell(100, 10, "Deduccion Pension (4%)", 1)
        pdf.cell(0, 10, f"- $ {result.pension:,.0f}", 1, ln=True)
        
        pdf.ln(5)
        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(100, 12, "NETO RECIBIDO", 0)
        pdf.cell(0, 12, f"$ {result.total_pagar:,.0f}", 0, ln=True)

        filename = f"colilla_{result.empleado.replace(' ', '_')}.pdf"
        pdf.output(filename)
        return filename

# =========================================================================
# VISTAS (INTERFACE GRÁFICA)
# =========================================================================
class LoginScreen(tb.Frame):
    """Pantalla de acceso inicial."""
    def __init__(self, master, on_success):
        super().__init__(master, padding=40)
        self.on_success = on_success

        tb.Label(self, text="NominaMaster Pro", font=("Helvetica", 22, "bold"), bootstyle=PRIMARY).pack(pady=20)
        tb.Label(self, text="Gestión Administrativa", font=("Helvetica", 10)).pack(pady=5)

        self.user = tb.Entry(self, font=("Helvetica", 12))
        self.user.pack(pady=10, fill=X)
        self.user.insert(0, "admin")

        self.password = tb.Entry(self, show="*", font=("Helvetica", 12))
        self.password.pack(pady=10, fill=X)
        self.password.insert(0, "admin")

        tb.Button(self, text="INICIAR SESIÓN", command=self.login, bootstyle=SUCCESS, width=20).pack(pady=30)

    def login(self):
        if self.user.get() == "admin" and self.password.get() == "admin":
            self.on_success()
        else:
            messagebox.showerror("Error", "Credenciales incorrectas")


class EmployeeView(tb.Frame):
    """Panel de administración de empleados (CRUD)."""

    def __init__(self, master):
        super().__init__(master, padding=20)
        
        # Titulo Seccion
        tb.Label(self, text="REGISTRO DE PERSONAL", font=("Helvetica", 14, "bold")).pack(anchor=W, pady=10)

        # Formulario
        form_frame = tb.Frame(self)
        form_frame.pack(fill=X, pady=10)
        
        self.vid = tb.StringVar()
        self.vnombre = tb.StringVar()
        self.vcargo = tb.StringVar()
        self.vsalario = tb.StringVar()
        self.vfecha = tb.StringVar()
        self.vhorario = tb.StringVar()

        # Layout del Formulario (Grid)
        fields = [
            ("ID / Documento:", self.vid, 0, 0),
            ("Nombre Completo:", self.vnombre, 0, 2),
            ("Cargo:", self.vcargo, 0, 4),
            ("Salario Base:", self.vsalario, 1, 0),
            ("Fecha Ingreso:", self.vfecha, 1, 2),
            ("Horario:", self.vhorario, 1, 4),
        ]

        for label, var, r, c in fields:
            tb.Label(form_frame, text=label).grid(row=r, column=c, padx=10, pady=10, sticky=W)
            tb.Entry(form_frame, textvariable=var, width=20).grid(row=r, column=c+1, padx=10, pady=10)

        # Botonera
        btn_frame = tb.Frame(self)
        btn_frame.pack(fill=X, pady=15)
        
        tb.Button(btn_frame, text="GUARDAR CAMBIOS", command=self.save_employee, bootstyle=SUCCESS).pack(side=LEFT, padx=5)
        tb.Button(btn_frame, text="LIMPIAR CAMPOS", command=self.clear_form, bootstyle=SECONDARY).pack(side=LEFT, padx=5)
        tb.Button(btn_frame, text="ELIMINAR EMPLEADO", command=self.delete_employee, bootstyle=DANGER).pack(side=RIGHT, padx=5)

        # Tabla (Treeview)
        self.tree = tb.Treeview(self, columns=("ID", "Nombre", "Cargo", "Salario", "Fecha", "Horario"), show="headings", bootstyle=INFO)
        for col in ("ID", "Nombre", "Cargo", "Salario", "Fecha", "Horario"):
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, anchor=CENTER, width=120)
            
        self.tree.pack(fill=BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.load_data()

    def load_data(self):
        """Recarga la información del Excel en la tabla."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.df = ExcelService.load_employees()
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=(
                row["ID"], row["Nombre"], row["Cargo"], 
                f"{row['Salario']:,.0f}", row["FechaIngreso"], row["Horario"]
            ))

    def on_select(self, event):
        """Mapea el clic en la tabla a las cajas de texto."""
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])['values']
            self.vid.set(item[0])
            self.vnombre.set(item[1])
            self.vcargo.set(item[2])
            # Limpiar formato de moneda para edición
            val_salario = str(item[3]).replace(",", "").replace(".", "")
            self.vsalario.set(val_salario)
            self.vfecha.set(item[4])
            self.vhorario.set(item[5])

    def clear_form(self):
        """Limpia todas las variables de entrada."""
        for var in [self.vid, self.vnombre, self.vcargo, self.vsalario, self.vfecha, self.vhorario]:
            var.set("")

    def save_employee(self):
        """Valida, limpia y guarda un empleado en el Excel."""
        emp_id = str(self.vid.get()).strip()
        nombre = self.vnombre.get().strip()
        salario_raw = self.vsalario.get().strip()

        # VALIDACIONES
        if not emp_id or not nombre or not salario_raw:
            messagebox.showwarning("Campos Vacíos", "ID, Nombre y Salario son obligatorios.")
            return
        
        if not DataValidator.is_valid_name(nombre):
            messagebox.showwarning("Nombre Inválido", "Por favor ingrese un nombre real.")
            return

        salario_clean = DataValidator.clean_numeric(salario_raw)
        if salario_clean <= 0:
            messagebox.showwarning("Salario Inválido", "El salario debe ser mayor a 0.")
            return

        self.df['ID'] = self.df['ID'].astype(str)
        idx = self.df.index[self.df["ID"] == emp_id].tolist()
        
        new_row = {
            "ID": emp_id,
            "Nombre": nombre,
            "Cargo": self.vcargo.get(),
            "Salario": salario_clean,
            "FechaIngreso": self.vfecha.get(),
            "Horario": self.vhorario.get()
        }

        if idx:
            for k, v in new_row.items():
                self.df.at[idx[0], k] = v
            messagebox.showinfo("Éxito", "Información actualizada correctamente.")
        else:
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
            messagebox.showinfo("Éxito", "Empleado registrado en el sistema.")

        ExcelService.save_employees(self.df)
        self.load_data()
        self.clear_form()

    def delete_employee(self):
        """Elimina el registro actualmente seleccionado."""
        emp_id = str(self.vid.get()).strip()
        if not emp_id:
            messagebox.showwarning("Atención", "Seleccione un empleado de la tabla.")
            return
            
        if messagebox.askyesno("Confirmar", f"¿Desea eliminar al ID {emp_id}?"):
            self.df['ID'] = self.df['ID'].astype(str)
            self.df = self.df[self.df["ID"] != emp_id]
            ExcelService.save_employees(self.df)
            self.load_data()
            self.clear_form()


class PayrollView(tb.Frame):
    """Panel de liquidación de pagos mensules."""

    def __init__(self, master):
        super().__init__(master, padding=20)
        
        self.df = ExcelService.load_employees()
        self.selected_employee = None
        self.last_payroll_result = None

        # Buscador superior
        self.search_group = tb.LabelFrame(self, text=" BUSCAR PERSONAL ")
        self.search_group.pack(fill=X, pady=10)
        search_frame = tb.Frame(self.search_group, padding=15)
        search_frame.pack(fill=X)
        
        tb.Label(search_frame, text="Documento ID:").pack(side=LEFT, padx=10)
        self.search_var = tb.StringVar()
        self.search_entry = tb.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=LEFT, padx=10)
        tb.Button(search_frame, text="CONSULTAR", command=self.load_employee, bootstyle=INFO).pack(side=LEFT, padx=10)
        
        # Detalles Resultado
        # Nota: Quitamos el parámetro padding del constructor para evitar TclError
        self.info_group = tb.LabelFrame(self, text=" INFORMACIÓN DE PAGO ")
        self.info_group.pack(fill=X, pady=10)
        # Aplicamos padding mediante configuración de empaquetado interna
        inner_info = tb.Frame(self.info_group, padding=15)
        inner_info.pack(fill=X)
        
        self.lbl_nombre = tb.Label(inner_info, text="Nombre: ---", font=("Helvetica", 12, "bold"))
        self.lbl_nombre.pack(anchor=W)
        self.lbl_salario = tb.Label(inner_info, text="Salario Base: $ 0", font=("Helvetica", 11))
        self.lbl_salario.pack(anchor=W, pady=5)

        # Entradas de cálculo
        calc_input = tb.Frame(inner_info)
        calc_input.pack(fill=X, pady=10)
        tb.Label(calc_input, text="Bonos / Horas Extra ($):").pack(side=LEFT)
        self.extra_var = tb.StringVar(value="0")
        tb.Entry(calc_input, textvariable=self.extra_var).pack(side=LEFT, padx=10)
        tb.Button(calc_input, text="CALCULAR NOMINA", command=self.calculate_payroll, bootstyle=PRIMARY).pack(side=LEFT)
        
        # Resultados finales
        self.res_display = tb.Frame(inner_info, padding=10, bootstyle=LIGHT)
        self.res_display.pack(fill=X, pady=10)

        self.lbl_dev = tb.Label(self.res_display, text="TOTAL DEVENGADO: $ 0")
        self.lbl_dev.pack(anchor=W)
        self.lbl_salud = tb.Label(self.res_display, text="SALUD (4%): $ 0", foreground="red")
        self.lbl_salud.pack(anchor=W)
        self.lbl_pen = tb.Label(self.res_display, text="PENSION (4%): $ 0", foreground="red")
        self.lbl_pen.pack(anchor=W)
        
        self.lbl_neto = tb.Label(self.res_display, text="NETO A PAGAR: $ 0", font=("Helvetica", 14, "bold"), foreground="green")
        self.lbl_neto.pack(anchor=W, pady=10)
        
        # Acciones
        self.btn_pdf = tb.Button(inner_info, text="IMPRIMIR COLILLA (PDF)", command=self.generate_pdf, bootstyle=SUCCESS, state=DISABLED)
        self.btn_pdf.pack(pady=10)

    def load_employee(self):
        """Busca el empleado en el DataFrame y carga su info base."""
        self.df = ExcelService.load_employees()
        self.df['ID'] = self.df['ID'].astype(str)
        emp_id = str(self.search_var.get()).strip()
        
        match = self.df[self.df["ID"] == emp_id] if emp_id else pd.DataFrame()

        if not match.empty:
            row = match.iloc[0]
            self.selected_employee = row
            self.lbl_nombre.config(text=f"Nombre: {row['Nombre']} | {row['Cargo']}")
            self.lbl_salario.config(text=f"Salario Base: $ {row['Salario']:,.0f}")
            self.btn_pdf.config(state=DISABLED)
        else:
            messagebox.showerror("No Encontrado", "El ID ingresado no existe en la base de datos.")

    def calculate_payroll(self):
        """Ejecuta los cálculos de ley."""
        if self.selected_employee is None:
            messagebox.showwarning("Atención", "Consulte un empleado primero.")
            return
            
        salario = float(self.selected_employee["Salario"])
        extra_clean = DataValidator.clean_numeric(self.extra_var.get())
        
        results = PayrollService.calcular_nomina(salario, extra_clean)
        
        # Actualizar UI
        self.lbl_dev.config(text=f"TOTAL DEVENGADO: $ {results['devengado']:,.0f}")
        self.lbl_salud.config(text=f"SALUD (4%): - $ {results['salud']:,.0f}")
        self.lbl_pen.config(text=f"PENSION (4%): - $ {results['pension']:,.0f}")
        self.lbl_neto.config(text=f"NETO A PAGAR: $ {results['neto']:,.0f}")
        
        self.last_payroll_result = PayrollResult(
            empleado=self.selected_employee["Nombre"],
            total_devengado=results["devengado"],
            salud=results["salud"],
            pension=results["pension"],
            total_pagar=results["neto"]
        )
        self.btn_pdf.config(state=NORMAL)

    def generate_pdf(self):
        """Dispara el servicio de PDF."""
        if self.last_payroll_result:
            try:
                fn = PDFService.generar_colilla(self.last_payroll_result)
                messagebox.showinfo("Reporte", f"Colilla guardada como: {fn}")
            except Exception as e:
                messagebox.showerror("Error PDF", f"No se pudo generar el archivo: {e}")


# =========================================================================
# APLICACIÓN PRINCIPAL
# =========================================================================
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SISTEMA DE GESTIÓN DE NÓMINA v2.0")
        self.root.geometry("1000x700")

        # Centrar ventana
        self.root.eval('tk::PlaceWindow . center')
        
        # Pantalla de Login
        self.login_manager = LoginScreen(self.root, self.on_login_success)
        self.login_manager.pack(expand=True)

    def on_login_success(self):
        """Callback al loguear con éxito."""
        self.login_manager.pack_forget()
        self.render_dashboard()
        
    def render_dashboard(self):
        """Crea el panel principal con pestañas."""
        main_container = tb.Frame(self.root)
        main_container.pack(fill=BOTH, expand=True)
        
        # Barra superior estética
        top_bar = tb.Frame(main_container, bootstyle=PRIMARY, height=50)
        top_bar.pack(fill=X)
        tb.Label(top_bar, text="PANEL DE CONTROL - NominaMaster Pro", foreground="white", font=("Helvetica", 12, "bold")).pack(pady=10)

        # Notebook
        self.tabs = tb.Notebook(main_container, bootstyle=INFO)
        self.tabs.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        # Vistas
        self.v_employees = EmployeeView(self.tabs)
        self.v_payroll = PayrollView(self.tabs)
        
        self.tabs.add(self.v_employees, text=" GESTIÓN DE EMPLEADOS ")
        self.tabs.add(self.v_payroll, text=" LIQUIDACIÓN DE NÓMINA ")


if __name__ == "__main__":
    # Iniciar ventana con tema oscuro moderno
    root_window = tb.Window(themename="darkly")
    MainApp(root_window)
    root_window.mainloop()