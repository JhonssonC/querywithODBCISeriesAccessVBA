# querywithODBCISeriesAccessVBA
Interfaz entre Excel(VBA) y conexión DSN ISeries Access ODBC Driver para obtención de datos a través de Query.


Prueba de Ejecución:

![Imgur](https://i.imgur.com/qdLrgvZ.gif)

*  Prerrequisitos (solo para windows 7 - 10): 

Office 2012-2019 (32/64 bits) 
Personal Communications iSeries Access para Windows

Instrucciones:

*  Abrir o crear un archivo Excel habilitado para macros.
*  Cree una tabla de contenidos a una hoja especifica llamada VAR.
Nota importante: Especificar Query Select en B1, el codigo de referencia para cada select es '[[CODE]]', y anexar tambien usuario y contraseña para la conexion en las celdas E1 y E2 respectivamente.

![Imgur2](https://i.imgur.com/JPWxF55.png)

SQL
```SQL
...
WHERE
C.CLCODE = '[[CODE]]'
...
```

*  En otra Hoja (puede ser Hoja1) construya en una hoja vacia la siguiente tabla poniendo especial atencion a las columnas especificadas en la hoja VAR (Columna A) en el paso anterior las columnas deben concordar con los encabezados, no textualmente pero si deben ser los datos que se especificaron el la hoja VAR.

![Imgur3](https://i.imgur.com/VWyjiod.png)

* Verifique que el acceso a datos del sistema (ODBC) tenga la siguiente configuracion. (ver codigo)

![Imgur4](https://i.imgur.com/iZ5JITV.png)

VBA
```VBA
USERNAME = Sheets("VAR").Range("E1")
PASSWORD = Sheets("VAR").Range("E2")
    
conn.ConnectionString = "dsn=SICOPUB-MAN;User Id=" & USERNAME & ";Password=" & PASSWORD & " ;"
```

Nota: Los datos objeto de busqueda son los codigos, estos se toman como referencia para ubicar el resto de datos en base al Query especificado en Hoja VAR aplicado a cada codigo.

Digitar codigos a buscar, seleccionar los codigos en la tabla y ejecutar la macro.

![Imgur](https://i.imgur.com/qdLrgvZ.gif)

Nota: La seleccion puede ser uno o varios elementos y soporta tambien elementos solo de un filtro especificado (previamente se deben filtrar los datos de la tabla en excel y unicamente ejecutara la macro a la seleccion sin considerar filas ocultas).
