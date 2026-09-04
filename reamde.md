# Exportador de Retenciones y Percepciones a TXT para DGR

Macro de Microsoft Excel en VBA para generar archivos `.txt` a partir de una tabla de retenciones y percepciones.

**NOTA:** Este macro está exclusivamente hecho para presentar en la declaración de IIBB de la página de la DGR en San Juan. Dicha página tiene su propio formato de archivo de presentación y puede difererir al cambiar la jurisdicción. 

**Importante:** Se usó Office 2007 para este macro debido a que es lo que tenía disponible en la oficina, por lo que algunas sentencias del código podrían no funcionar en otras versiones. Es importante tener esto en cuenta a la hora de ejecutarlo, ya que puede largar errores relacionados con la sintaxis.

La macro genera dos archivos separados:

* `DGR_Retenciones.txt`
* `DGR_Percepciones.txt`

Los archivos se generan en la misma carpeta donde está guardado el archivo de Excel que contiene la macro.

## Requisitos

* Microsoft Excel con soporte para macros VBA. 
* Un archivo de Excel con la estructura de tabla indicada en este documento.
* El libro debe estar guardado antes de ejecutar la macro.

## Estructura de la tabla

La macro toma los datos de la **hoja activa** y espera que la tabla tenga encabezados en la fila 1 y datos a partir de la fila 2.

Las columnas deben estar ubicadas de la siguiente manera:

| Columna | Campo        | Descripción                                                                               |
| ------- | ------------ | ----------------------------------------------------------------------------------------- |
| A       | Tipo         | Tipo de comprobante. Se utiliza para determinar si corresponde Factura o Nota de Crédito. |
| B       | Denominación | Nombre o razón social del sujeto.                                                         |
| C       | CUIT         | CUIT del sujeto informado.                                                                |
| D       | Concepto     | Debe contener "Retención" o "Percepción".                                                 |
| E       | Orden        | Número de orden utilizado para generar el número de comprobante de las percepciones.      |
| F       | Neto         | Importe neto sujeto a retención/percepción.                                               |
| G       | IB           | Importe de la retención o percepción de Ingresos Brutos.                                  |

### Ejemplo de estructura

| A - Tipo        | B - Denominación   | C - CUIT    | D - Concepto | E - Orden |  F - Neto |  G - IB |
| --------------- | ------------------ | ----------- | ------------ | --------: | --------: | ------: |
| Factura         | EMPRESA EJEMPLO SA | 30700000001 | Percepción   |         1 | 100000.00 | 3600.00 |
| Factura         | COMERCIO EJEMPLO   | 20300000002 | Retención    |         2 |  50000.00 | 1250.00 |
| Nota de Crédito | EMPRESA EJEMPLO SA | 30700000001 | Percepción   |         3 | -20000.00 | -720.00 |

No es necesario que los encabezados tengan exactamente esos nombres, pero **las columnas deben respetar las posiciones A a G indicadas**.

## Consideraciones sobre los datos

### Tipo de comprobante

La macro determina automáticamente el tipo de comprobante:

* Si la columna A contiene `Factura`, utiliza `01`.
* Si contiene `Nota` o `Crédito`, utiliza `03`.
* Si no encuentra ninguno de esos valores, utiliza `01` por defecto.

### Concepto

La columna D determina a qué archivo se exporta cada registro:

* Si contiene `Retención`, se agrega a `DGR_Retenciones.txt`.
* Si contiene `Percepción`, se agrega a `DGR_Percepciones.txt`.

Los registros cuyo concepto no contenga ninguno de esos términos no se exportan.

###Alícuotas

La versión actual de la macro está configurada para las alícuotas utilizadas en el régimen de Ingresos Brutos de la Provincia de San Juan:

Percepciones: 3,60 %
Retenciones: 2,50 %

Estas alícuotas están definidas directamente en el código VBA. Verificar que correspondan al régimen y período que se desea informar antes de utilizar la macro.


> Estos valores se encuentran definidos como constantes en el código VBA::

```vb
Const ALICUOTA_PERCEPCION As String = "3.60"
Const ALICUOTA_RETENCION As String = "2.50"
```

Si la macro se utiliza para otra jurisdicción o régimen con alícuotas diferentes, deben modificarse estos valores antes de ejecutar el proceso.

Se recomienda verificar siempre las alícuotas vigentes para el período que se desea informar.

### Valores negativos

Los importes mantienen el signo original de la planilla.

Por ejemplo:

```text
20000.00
-20000.00
```

Esto permite conservar valores negativos, como los correspondientes a notas de crédito, restricción que impone el sitio web de la DGR San Juan.

## Ejecución

1. Abrir el archivo de Excel 
2. Abrir VBA en las opciones para desarolladores e importar el archivo .bas
3. Verificar que la hoja activa tenga la estructura indicada.
4. Guardar el archivo de Excel.
5. Seleccionar la hoja con los datos.
6. Ejecutar `GenerarTxtDGR()`.
7. Ingresar la fecha del período cuando la macro lo solicite.
8. La macro generará los dos archivos `.txt` en la misma carpeta del archivo de Excel.

## Archivos generados

### Retenciones

`DGR_Retenciones.txt`

Contiene los registros correspondientes a retenciones.

El formato generado es:

```text
CUIT    DENOMINACION    FECHA    TIPO_COMP    NETO    ALICUOTA    IB    0.00
```

### Percepciones

`DGR_Percepciones.txt`

Contiene los registros correspondientes a percepciones.

El formato generado es:

```text
CUIT    DENOMINACION    FECHA    TIPO_COMP    NUM_COMP    NETO    ALICUOTA    IB    0.00
```

**NOTA:** Los campos están separados por tabulaciones.

## Número de comprobante en percepciones

Para las percepciones, el número de comprobante se genera utilizando:

* El número de orden de la columna E.
* La fecha ingresada al ejecutar la macro.

El número de orden se completa con ceros hasta alcanzar cinco posiciones.

Por ejemplo:

```text
Orden: 25
Fecha: 04/09/2026
```

Genera:

```text
00025-04092026
```

## Notas

* La macro trabaja sobre la **hoja activa**.
* Los datos comienzan en la fila 2.
* La última fila se determina utilizando la columna C (CUIT).
* Los archivos existentes con el mismo nombre son sobrescritos al ejecutar la macro.
* Los archivos se generan utilizando tabulaciones como separadores.
* La macro no almacena datos de clientes dentro del código: los datos se toman directamente de la planilla.

## Estructura recomendada del repositorio


Se recomienda no subir al repositorio planillas reales utilizadas para procesar información de clientes, especialmente si contienen CUIT, denominaciones, importes u otros datos identificables.

