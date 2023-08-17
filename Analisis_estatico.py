import pandas as pd
from fpdf import FPDF
import comtypes.client
#OBTENER EL COEFICIENTE SÍSMICO
#Conectar a etabs y obtener periodos
def connect_to_etabs():
    ETABSobject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    SapModel = ETABSobject.SapModel
    return SapModel
#Obtener periodos de etabs
def periodos_etabs(SapModel):
    SapModel.SetPresentUnits(12)
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetCaseSelectedForOutput("Modal")
    data = SapModel.Results.ModalParticipatingMassRatios()
    Modal = pd.DataFrame(data[1:17],index= ['LoadCase', 'StepType', 'StepNum', 
                                            'Period', 'Ux', 'Uy', 'Uz','SumUx', 
                                            'SumUy', 'SumUz', 'Rx', 'Ry', 'Rz', 
                                            'SumRx', 'SumRy', 'SumRz']).transpose()
    
    data_necesitada = ['LoadCase', 'Period',"Ux", "Uy", "Rz"]
    Modal = Modal[data_necesitada]
    Modal_filtrado = Modal[Modal['LoadCase'] == "Modal"]
    modal_x_indice = Modal_filtrado[Modal_filtrado.Ux == max(Modal_filtrado.Ux)].index
    modal_y_indice = Modal_filtrado[Modal_filtrado.Uy == max(Modal_filtrado.Uy)].index
    modal_z_indice = Modal_filtrado[Modal_filtrado.Rz == max(Modal_filtrado.Rz)].index
    T_x = Modal_filtrado.Period[modal_x_indice[0]]
    T_y = Modal_filtrado.Period[modal_y_indice[0]]
    T_z = Modal_filtrado.Period[modal_z_indice[0]]
    T_fundamental = Modal_filtrado.Period[0]
    return T_x, T_y, T_z, T_fundamental

def analisis(T_x, T_y):
    #VALORES PARA LAS TABLAS
    valores_z = {"Z4":0.45,"Z3":0.35,"Z2":0.25,"Z1":0.1}
    valores_u = {"A":1.5,"B":1.3, "C":1}
    valores_s = {"S0": [0.8,0.8,0.8,0.8],
                "S1": [1,1,1,1],
                "S2": [1.05,1.15,1.20,1.60],
                "S3": [1.10,1.20,1.40,2]}
    valores_tp_tl = {"S0": [0.3,3],
                    "S1": [0.4,2.5],
                    "S2": [0.6,2],
                    "S3": [1,1.6]}

    #TABLAS E0.30 - PERÚ
    tabla_s = pd.DataFrame(valores_s, index=["Z4","Z3","Z2","Z1"])
    tabla_tp_tl = pd.DataFrame(valores_tp_tl, index = ["TP","TL"])

    #DATOS INGRESADOS POR EL USUARIO
    Z_nombre = input("Ingresa la zona sismica de la edificación: ").upper()
    U_nombre = input("Ingresa la categoría de la edificación: ").upper()
    S0 = input("Ingresa el perfil de suelo: ").upper()
    Ro = int(input("R asumido: "))

    #OBTENCIÓN DE VALORES
    TP, TL = tabla_tp_tl.loc["TP",S0],tabla_tp_tl.loc["TL",S0]
    Z = valores_z[Z_nombre]
    U = valores_u[U_nombre]
    Cx = 2.5 if T_x < TP else (2.5 * (TP / T_x) if TP < T_x < TL 
                               else 2.5 * (TP * TL / T_x**2))
    Cy = 2.5 if T_y < TP else (2.5 * (TP / T_y) if TP < T_y < TL else 2.5 * (TP * TL / T_y**2))
    S = tabla_s.loc[Z_nombre,S0]
    ZUCS_R_X = (Z*U*Cx*S/Ro)
    ZUCS_R_Y = (Z*U*Cy*S/Ro)
    T_values = [T_x, T_y]
    K_values = [1 if i <= 0.5 else 0.75 + 0.5 * i if (0.75 + 0.5 * i) <= 2 else 0 for i in T_values]
    resultado = [f"Factor Z = {Z}g",
                f"Factor U = {U}",
                f"Factor Cx = {Cx}",
                f"Factor Cy = {Cy}",
                f"Factor S = {S}",
                f"Factor Ro = {Ro}",
                f"TP = {TP}s, TL = {TL}s",
                f"Tx ={round(T_x,4)}s, Ty = {round(T_y,4)}s",
                f"ZUCxS/R = {round(ZUCS_R_X,6)}",
                f"ZUCyS/R= {round(ZUCS_R_Y,6)}",
                f"Kx= {K_values[0]}",
                f"Ky= {K_values[1]}"]
    valores = [Z,U,Cx,Cy,S,Ro,TP,TL,ZUCS_R_X,ZUCS_R_Y,K_values[0],K_values[1] ]
    #MOSTRAR RESULTADOS
    print("----------VALORES----------")
    print('\n'.join(resultado))
    return resultado,valores

def pdf(resultado):
    doc = FPDF(orientation='P', unit ='mm', format='A4')
    doc.add_page()
    doc.set_font('Courier', 'BU', 14)
    #Título
    texto_ancho = doc.get_string_width("ANÁLISIS ESTÁTICO - E030 2018")
    pos_x = (doc.w - texto_ancho) / 2
    doc.text(x=pos_x, y=20, txt="ANÁLISIS ESTÁTICO - E030 2018")
    #Cuerpo
    doc.set_font('Courier', '', 12)
    h = 8
    for i,j in zip(range(20+h,(20+h+(h*len(resultado))),h),resultado):
        doc.text(x=24, y= i, txt=j)
    #GUARDAR PDF
    doc.output(f'Y:/Downloads/Documents/ANÁLISIS_ESTÁTICO.pdf')

if __name__ == "__main__":
    SapModel = connect_to_etabs()
    T_x, T_y,T_z, T_fundamental= periodos_etabs(SapModel)
    resultado, valores = analisis(T_x,T_y)
    pdf(resultado)

    

