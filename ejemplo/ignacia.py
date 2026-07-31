"""Las decisiones de identidad de Ignacia Fuentes, en un solo sitio.

Ignacia es una persona ficticia. Su sistema esta publicado en el repositorio
hermano, https://github.com/Carlos-Padilla-Bravo/identidad-personal, donde el
mismo manual sale en HTML y en PDF. Aqui sale en Word, con los mismos valores.

Este archivo existe para separar lo que decide la persona de lo que compone la
libreria. `manual.py` y `figura.py` lo importan los dos, asi que el color se
escribe una vez. Es el reparto que conviene copiar: ooxmlkit no sabe nada de
Ignacia, y este modulo no sabe nada de OOXML.

MANDA EL MANUAL DE ALLA, Y NO SE MUEVE. Sus valores estan cerrados y publicados,
asi que este archivo es una copia y nunca al reves. El riesgo no es que aquel
cambie: es editar ESTE y dejarlo diciendo algo que el manual de ella contradice.
Paso al escribirlo, dos veces, y no lo delato ninguna prueba sino mirar la tira
de muestra del README de alla. Lo que ella declara, para no tener que ir a
buscarlo:

    Cuadro 3, familias y usos    Fraunces      Titulos, firma, citas
                                 Nunito Sans   Cuerpo, cuadros, apoyos
                                 Space Mono    Cifras, tokens, tablas de aporte
    Figura 3, proporcion de uso  90 % neutros, 8 % primario, 2 % senal

Al tocar este archivo, contrastar contra esas cuatro lineas. Si alguna vez el
manual de ella si cambiara, esto pasa a estar mal en silencio y no hay nada que
lo detecte.
"""

# Los diez roles que ooxmlkit espera, con los valores de ella. Su sistema declara
# dos roles de texto; los otros dos se derivaron por calculo y esta dicho abajo
# cual y por que.
PALETA = {
    "texto":        "#2A2622",  # su tinta calida
    "texto_fuerte": "#3A332C",
    "texto_sec":    "#6E6459",
    "texto_ter":    "#796E62",  # derivado, ver NOTA_TEXTO_TER
    "primario":     "#3C6E54",  # el verde
    "primario_int": "#2E5943",
    "senal":        "#AA4C2C",  # la terracota, racionada
    "fondo":        "#FFFFFF",  # el papel, ver NOTA_HOJA
    "superficie":   "#F1E9DB",  # su crema, como bloque tramado
    "borde":        "#E5DBCB",
}

# Sus tres familias, todas de licencia OFL. Word compone con lo que este
# instalado y sustituye sin avisar, asi que si el documento sale con otra letra
# es que faltan: se instalan desde Google Fonts.
#
# OJO CON EL REPARTO. Las constantes de ooxmlkit se llaman SERIF, SANS y COND,
# pero lo que nombran es un ENCARGO y no una clasificacion: SERIF es la letra del
# cuerpo y SANS la de los titulos. En el sistema de Ignacia esos encargos van al
# reves de lo que sugiere el nombre, porque su serif titula y su sans es la de
# leer. Asignar por clasificacion en vez de por encargo produce un manual que
# contradice lo que el mismo declara en su cuadro de familias.
CUERPO = "Nunito Sans"      # "Cuerpo, cuadros, apoyos", dice su manual
TITULOS = "Fraunces"        # "Titulos, firma, citas"
CIFRAS = "Space Mono"       # "Cifras, tokens, tablas de aporte"

SERIF = CUERPO              # el encargo de SERIF en ooxmlkit es el cuerpo
SANS = TITULOS              # y el de SANS, los titulos
MONO = CIFRAS
COND = CUERPO               # su sistema no tiene condensada: repite la de leer

# Cuanto de cada cosa en una pieza. El reparto lo declara su manual y esta
# copiado, no elegido: iba en 88/9/3 hasta que se contrasto contra el original.
PROPORCION = [("Neutros", 90), ("Primario", 8), ("Señal", 2)]

# Estas dos cadenas se imprimen dentro del documento, asi que van con tildes.
# Los comentarios y las cadenas de codigo de este repositorio van sin ellas.
NOTA_TEXTO_TER = (
    "Su sistema declara dos roles de texto y un documento compuesto necesita "
    "cuatro. El más claro se derivó por cálculo y no a ojo: mismo matiz y misma "
    "saturación que su texto secundario, y la claridad más alta que todavía "
    "alcanza AA sobre la hoja.")

NOTA_HOJA = (
    "En pantalla su fondo es crema. El papel es blanco, así que la crema pasa a "
    "ser superficie y los contrastes se recalculan contra la hoja: un valor está "
    "calibrado contra una superficie y deja de valer cuando la superficie se "
    "mueve.")
