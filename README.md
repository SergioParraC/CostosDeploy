# CostosDeploy
Pasa a un documento de Word las cantidades de obra y los Análisis de Precios Unitarios que están en Excel
---------------------------------------------------------------------------------------------------------

##Como usar el código

1. Se debe tener los capítulos dividas por hojas
2. Dentro de las hojas deben estar los items, en una tabla independiente, use la hoja original para crear sus items con cantidades de obra calculadas a mano
3. Asigne en la esquina superior izquierda en el siguiente orden el codigo de la tabla:
    1. Numero del capitulo
    2. Numero del tipo de elemento calculado (ej, Apto tipo 2=2, Cubierta=6)
    3. Numero de ítem dentro del capitulo
4. En la hoja "EDS" ponga:
    1. En la fila del capitulo columna A el codigo asignado al capitulo
    2. En la fila del ítem columna B el código del ítem
    3. En la columna C un codigo unico con el que quiera identificar su item dentro del trabajo
    4. En la fila 1 y las columnas D los elementos que usted calculo con su respectivo codigo, asignados en el paso 3
5. En la hoja "APU" ponga sus analisis de precios unitarios calculados para cada item en el orden especificado en el ejemplo. 
6. Añada el script en un modulo de su archivo excel y ejecute el codigo

##Salida del programa

Un archivo en Word con las siguientes características:
* Capítulos se asignan como título 1
* Ítem se asignan como título 2
* Elementos se asignan como título 3

Configurar los titulos a gusto del usuario y terminar el documento

---------------------------------------------------------------------------------------------------------

Este codigo fue realizado por hobbie y hacer mi trabajo como Ingeniero de Costos y presupuestos mas agradable, divertido y agil :)