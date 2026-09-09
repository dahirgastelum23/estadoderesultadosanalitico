# Estado de Resultados Analítico

Sistema desarrollado en Python para calcular y generar un Estado de Resultados Analítico de una empresa, con exportación automática a Excel.

## Descripción

Este proyecto permite ingresar información financiera de una empresa y calcular automáticamente los principales resultados contables, incluyendo:

- Ventas netas
- Compras totales
- Compras netas
- Mercancía disponible
- Costo de ventas
- Utilidad bruta
- Gastos de operación
- Utilidad de operación
- Resultado antes de impuestos
- PTU
- ISR
- Utilidad neta o pérdida neta

El sistema también genera un archivo de Excel con el estado de resultados y sus respectivas fórmulas.

## Tecnologías utilizadas

- Python
- OpenPyXL
- Microsoft Excel

## Funcionalidades

### Captura de información

El programa solicita datos como:

- Ventas
- Devoluciones, rebajas, descuentos y bonificaciones sobre ventas
- Compras
- Gastos sobre compras
- Devoluciones, rebajas, descuentos y bonificaciones sobre compras
- Inventarios
- Gastos de venta
- Gastos de administración
- Depreciaciones
- Productos financieros
- Otros productos
- Gastos financieros
- Otros gastos

### Cálculos automáticos

El programa realiza automáticamente los cálculos necesarios para obtener el resultado financiero de la empresa.

Si el resultado final es positivo, se muestra:

**UTILIDAD NETA**

Si el resultado final es negativo:

**PÉRDIDA NETA**

Si el resultado es exactamente cero:

**RESULTADO NETO**

## Exportación a Excel

Al finalizar, el programa genera automáticamente un archivo `.xlsx`.

Los archivos se guardan utilizando nombres consecutivos para evitar sobrescribir reportes anteriores:

```text
estado_resultados.xlsx
estado_resultados_2.xlsx
estado_resultados_3.xlsx
