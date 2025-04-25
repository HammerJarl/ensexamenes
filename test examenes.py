import random

# generador de examenes, proyecto realizado para la ENS de bolivia, con el fin de tener una herramienta que les permita generar examenes.
banco_preguntas = []

#Función para generar opciones incorrectas
def generar_opciones_incorrectas(pregunta, respuesta_correcta):
    opciones_incorrectas = []

    if "capital" in pregunta.lower():
        posibles = ["Cochabamba", "Santa Cruz", "Oruro", "Riberalta"]
        opciones_incorrectas = random.sample([x for x in posibles if x != respuesta_correcta],4)
    elif "año" in pregunta.lower():
        #Si es sobre años, generamos años cercanos pero incorrectos
        try:
            año_base = int (respuesta_correcta)
            opciones_incorrectas = [str(año_base + random.randint (-10, 10)) for _ in range (4)]
            while respuesta_correcta in opciones_incorrectas:
                opciones_incorrectas = [str(año_base + random.randint (-10, 10)) for _ in range (4)]
        except ValueError:
            opciones_incorrectas = ["1999", "2000", "2001", "2002"]
    else:
        opciones_incorrectas = [f"{i+1}" for i in range(4)]

    return opciones_incorrectas
#admin
def ingresar_preguntas():
    print("=== MODO ADMINISTRADOR: INGRESO DE PREGUNTAS ===")
    while len(banco_preguntas) < 100:
        pregunta = input("Ingrese la pregunta (o escriba 'salir' para terminar): ")
        if pregunta.lower() =="salir":
            break
        respuesta_correcta = input ("Ingrese la respuesta correcta: ")

        opciones_incorrectas = generar_opciones_incorrectas (pregunta, respuesta_correcta)

        opciones = opciones_incorrectas + [respuesta_correcta]
        random.shuffle(opciones) #mezcla
        letra_correcta = chr(97 + opciones.index(respuesta_correcta)) # a, b, c, d, e

        banco_preguntas.append({
                "pregunta": pregunta,
                "opciones": opciones,
                "respuesta": letra_correcta
        })
        print (f"Total de preguntas ingresadas: {len(banco_preguntas)}")
    
def generar_examen(num_preguntas):
    if num_preguntas > len(banco_preguntas):
        print (f"No hay suficientes pregunas. El máximo disponible es: {len(banco_preguntas)}")
        return None
    if num_preguntas <=0:
        print ("El número de preguntas debe ser mayor a 0")
        return None
        
    #seleccionar y randomizar preguntas
    preguntas_seleccionadas = random.sample(banco_preguntas, num_preguntas)
    examen = []
    for i, item in enumerate(preguntas_seleccionadas, 1):
        opciones_random = item["opciones"].copy()
        random.shuffle(opciones_random) #random
        examen.append({
            "pregunta": f"Pregunta {i}: {item['pregunta']}",
            "opciones": opciones_random,
            "repuesta": chr(97 + opciones_random.index(item['opciones'][ord(item['respuesta']) - 97]))
         })
    return examen

def mostrar_examen (examen, version):
    print(f"\n=== EXAMEN VERSION {version} ===")
    for item in examen:
        print(item["pregunta"])
        for j, opcion in enumerate (item["opciones"]):
            print(f"{chr(97 + j)}) {opcion}")
        print ()
    
#menu ususario final
def menu_usuario():
    if not banco_preguntas:
        print("No hay preguntas disponibles. El administrador debe ingresarlas primero.")
        return
        
    print("\n=== MODO USUARIO: GENERAR EXAMEN ===")
    try:
        num_preguntas = int (input("cuantas preguntas desea en el examen? (màximo 100): "))
        if num_preguntas >100:
            print("El limite maximo de preguntas es 100")
            return
        #generar 2 examenes
        examen1 = generar_examen(num_preguntas)
        examen2 = generar_examen(num_preguntas)

        if examen1 and examen2:
            mostrar_examen(examen1, "A")
            mostrar_examen(examen2, "B")
    except ValueError:
        print("Por favor, ingresa un número valido")

#main
if __name__ == "__main__":
    # admin ingresa preguntas
    ingresar_preguntas()

    #usuario final genera examen
    menu_usuario()