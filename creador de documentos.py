import re
from datetime import datetime, timedelta
from Multiherramienta import *
import tkinter as tk
from tkinter import ttk
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

def generar_pdf():
    # Obtener valores de las entradas de texto
    cantidad_moldes = entry_cantidad_moldes.get()
    valor_accesorios = entry_valor_accesorios.get()
    valor_accesorios_letras = convertir_numero_a_texto(float(valor_accesorios))
    fraccionamiento = entry_fraccionamiento.get()
    contratante = entry_contratante.get()
    contratista = entry_contratista.get()
    domicilio_fraccionamiento = entry_domicilio_fraccionamiento.get()
    prototipo = entry_prototipo.get()
    valor_molde = entry_valor_molde.get()
    valor_molde_letras = convertir_numero_a_texto(float(valor_molde))
    
    # Obtener fecha de hoy y calcular fecha de vencimiento (+360 días)
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    fecha_vencimiento = (datetime.now() + timedelta(days=360)).strftime("%d/%m/%Y")

    # Crear el archivo PDF
    ruta = 'C:\\Users\\fprado\\OneDrive - Servicios Administrativos Javer, S.A. DE C.V\\Escritorio\\DOCUMENTOS\\CONTRATOS\\pagare.pdf'
    pdf = SimpleDocTemplate("pagare.pdf", pagesize=LETTER)
    elements = []
    width, height = LETTER

    # Definir estilos
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle("Titulo", parent=styles["Title"], alignment=1, fontSize=14)
    estilo_negritas = ParagraphStyle("Negritas", parent=styles["BodyText"], fontName="Helvetica-Bold")
    estilo_izquierda = styles["BodyText"]
    estilo_justificado = ParagraphStyle("Justificado", parent=styles["BodyText"], alignment=4)
    estilo_derecha = ParagraphStyle("Derecha", parent=styles["BodyText"], alignment=2)
    
    # Página 1
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(widht / 2, height - 50, "PAGARÉ")
    
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(100, height - 100, "Por concepto de ACCESORIOS:")
    pdf.setFont("Helvetica", 10)
    texto_parrafo_1 = (
        f"Cantidad entregada: {cantidad_moldes} paquete(s) de accesorios\n"
        f"Valor accesorios: {valor_accesorios}\n"
        f"Valor accesorios: {valor_accesorios}\n"
        f"Fraccionamiento: {fraccionamiento}")
    
    pdf.drawString(100, height - 200, f"Por este pagaré causal me obligo a pagar incondicionalmente a la orden de {contratante},")
    pdf.drawString(100, height - 220, f"la cantidad de {valor_accesorios} ({valor_accesorios_letras}) correspondiente al 100% del valor del(los)")
    pdf.drawString(100, height - 240, f"paquete(s) de accesorios entregado(s), teniendo como fecha de vencimiento el día {fecha_vencimiento},")
    pdf.drawString(100, height - 260, f"en su domicilio ubicado en {domicilio_fraccionamiento}.")
    
    pdf.drawString(100, height - 300, f"El presente pagaré es \"NO NEGOCIABLE\" y se deriva del contrato de obra de precio tiempo")
    pdf.drawString(100, height - 320, f"determinado, formalizado el 03 de Julio del 2023, en el que la \"CONTRATANTE\" es {contratante},")
    pdf.drawString(100, height - 340, f"y el \"CONTRATISTA\" {contratista}.")
    
    pdf.drawString(100, height - 380, f"Para el caso de que el importe que ampara el presente pagaré no sea liquidado puntualmente,")
    pdf.drawString(100, height - 400, f"causará intereses moratorios a razón de aplicar sobre el importe no pagado una Tasa de Interés")
    pdf.drawString(100, height - 420, f"del 12% anual. Los intereses se causarán y calcularán sobre la base de 360 días, y se cobrarán")
    pdf.drawString(100, height - 440, f"por los días realmente transcurridos desde la fecha del vencimiento y hasta la fecha de la")
    pdf.drawString(100, height - 460, f"liquidación del pago de que se trate.")
    
    pdf.drawString(100, height - 500, f"{domicilio_fraccionamiento}, {fecha_hoy}")
    
    pdf.showPage()
    
    # Página 2
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(100, height - 50, "PAGARÉ")
    
    pdf.setFont("Helvetica", 10)
    pdf.drawString(100, height - 100, "Por concepto de MOLDE:")
    pdf.drawString(100, height - 120, f"Cantidad entregada: {cantidad_moldes} molde(s)")
    pdf.drawString(100, height - 140, f"Tipo(s) de molde(s): {prototipo}")
    pdf.drawString(100, height - 160, f"Valor molde(s): {valor_molde}")
    pdf.drawString(100, height - 180, f"Fraccionamiento: {fraccionamiento}")
    
    pdf.drawString(100, height - 220, f"Por este pagaré causal me obligo a pagar incondicionalmente a la orden de {contratante},")
    pdf.drawString(100, height - 240, f"la cantidad de {valor_molde} ({valor_molde_letras}) correspondiente al 20% del valor del(los)")
    pdf.drawString(100, height - 260, f"molde(s) entregado(s), teniendo como fecha de vencimiento el día {fecha_vencimiento},")
    pdf.drawString(100, height - 280, f"en su domicilio ubicado en {domicilio_fraccionamiento}.")
    
    pdf.drawString(100, height - 320, f"El presente pagaré es \"NO NEGOCIABLE\" y se deriva del contrato de obra de precio tiempo")
    pdf.drawString(100, height - 340, f"determinado, formalizado el 03 de Julio del 2023, en el que la \"CONTRATANTE\" es {contratante},")
    pdf.drawString(100, height - 360, f"y el \"CONTRATISTA\" {contratista}.")
    
    pdf.drawString(100, height - 400, f"Para el caso de que el importe que ampara el presente pagaré no sea liquidado puntualmente,")
    pdf.drawString(100, height - 420, f"causará intereses moratorios a razón de aplicar sobre el importe no pagado una Tasa de Interés")
    pdf.drawString(100, height - 440, f"del 12% anual. Los intereses se causarán y calcularán sobre la base de 360 días, y se cobrarán")
    pdf.drawString(100, height - 460, f"por los días realmente transcurridos desde la fecha del vencimiento y hasta la fecha de la")
    pdf.drawString(100, height - 480, f"liquidación del pago de que se trate.")
    
    pdf.drawString(100, height - 520, f"{domicilio_fraccionamiento}, {fecha_hoy}")
    
    pdf.save()
    
    print("PDF generado exitosamente como 'pagare.pdf'.")



class PDFBuilder:
    def __init__(self, filename="output.pdf"):
        self.filename = filename
        self.story = []
        self.styles = getSampleStyleSheet()
        self.paragraphs = []

    def add_paragraph(self, text, style_name="BodyText"):
        """
        Adds a paragraph with text that may contain variables in the format {variable}.
        :param text: Paragraph text with variables in the format {variable}
        :param style_name: Style name (optional)
        """
        style = self.styles[style_name]
        variables = re.findall(r"\{(\w+)\}", text)  # Find all variables in the text
        self.paragraphs.append({"text": text, "style": style, "variables": variables})

    def set_variables(self, variable_values):
        """
        Assigns values to the paragraph variables.
        :param variable_values: Dictionary with values for each variable
        """
        for paragraph in self.paragraphs:
            text = paragraph["text"]
            for var in paragraph["variables"]:
                if var in variable_values:
                    # Replace the variable in the text
                    text = text.replace(f"{{{var}}}", str(variable_values[var]))
                else:
                    print(f"Warning: Variable '{var}' not found in provided values.")
            # Add the final paragraph to the story
            self.story.append(Paragraph(text, paragraph["style"]))

    def build_pdf(self):
        """
        Generates the final PDF with all paragraphs in `self.story`.
        """
        doc = SimpleDocTemplate(self.filename)
        doc.build(self.story)
        print(f"PDF generated at {self.filename}")
        # Clear the story for future builds if needed
        self.story = []


class VariableCapturer:
    def __init__(self, variables):
        """
        Initializes the variable capture window.
        :param variables: List of variable names to capture.
        """
        self.variables = variables
        self.values = {}  # Dictionary to store variable values
        self.root = tk.Tk()
        self.root.title("Capture Variables")
        
        # Interface configuration
        self.create_widgets()
        self.root.mainloop()
        
    def create_widgets(self):
        """
        Creates interface widgets, including input fields for each variable.
        """
        self.entries = {}
        
        # Dynamic creation of input fields based on variables
        for i, var in enumerate(self.variables):
            label = ttk.Label(self.root, text=f"{var}:")
            label.grid(row=i, column=0, padx=5, pady=5, sticky="w")
            
            entry = ttk.Entry(self.root)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.entries[var] = entry

        # Button to confirm and capture values
        submit_button = ttk.Button(self.root, text="Capture Values", command=self.capture_values)
        submit_button.grid(row=len(self.variables), column=0, columnspan=2, pady=10)
        
    def capture_values(self):
        """
        Captures entered values and saves them in `self.values`.
        Then closes the window.
        """
        for var, entry in self.entries.items():
            self.values[var] = entry.get()
        
        self.root.destroy()  # Closes the window

    def get_values(self):
        """
        Returns the captured values.
        """
        return self.values

def documento_adendum():
    
    


''' Configuración de la interfaz con Tkinter
root = tk.Tk()
root.title("Generador de Pagaré")

labels_texts = [
    "Cantidad Moldes", "Valor Accesorios", "Fraccionamiento",
    "Contratante", "Contratista", "Domicilio Fraccionamiento",
    "Prototipo", "Valor Molde"
]

entries = {}
for text in labels_texts:
    label = ttk.Label(root, text=text)
    label.pack()
    entry = ttk.Entry(root, width=40)
    entry.pack()
    entries[text] = entry

entry_cantidad_moldes = entries["Cantidad Moldes"]
entry_valor_accesorios = entries["Valor Accesorios"]
entry_fraccionamiento = entries["Fraccionamiento"]
entry_contratante = entries["Contratante"]
entry_contratista = entries["Contratista"]
entry_domicilio_fraccionamiento = entries["Domicilio Fraccionamiento"]
entry_prototipo = entries["Prototipo"]
entry_valor_molde = entries["Valor Molde"]

# Botón para generar el PDF
button = ttk.Button(root, text="Generar PDF", command=generar_pdf)
button.pack()

root.mainloop()
'''