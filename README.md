# Análisis de Ventas por Región y Producto (Excel)

Este proyecto muestra un análisis completo de ventas utilizando exclusivamente **Microsoft Excel**. Incluye automatización mediante fórmulas, panel de control dinámico, uso de tablas dinámicas, formato condicional y una macro para cálculo de precios por unidad.

## 📊 Funcionalidades principales

- **Panel de control (C.mando)** con selección de año para:
  - Contar el número de regiones por producto
  - Sumar unidades vendidas por producto
  - Sumar ventas (€) por producto
- **Subtotales por región y por vendedor**
- **Tablas dinámicas** para análisis por región, producto, vendedor y año
- **Gráfico dinámico** incluido
- **Formato condicional** aplicado a las ventas para resaltar importes extremos
- **Macro "PRECIO/UD"** para calcular automáticamente el precio unitario

## 📁 Estructura del repositorio

```
excel-analisis-ventas/
├── informe-vntas-excel.xlsx # Archivo principal con todos los cálculos
├── macro.xlsx # Archivo que contiene la macro grabada
├── README.md # Descripción del proyecto
└── Tablas_Dinámicas_Gráfico.pdf # Vista previa del panel de Excel
```

## 🛠 Herramientas y funciones usadas

- Funciones avanzadas de Excel:
  - `SUMAR.SI.CONJUNTO`, `CONTAR.UNICOS`, `FILTRAR`, `VALOR`, `BUSCARV`
- Tablas dinámicas con segmentación por año, región y producto
- Gráfico dinámico generado a partir de tablas dinámicas
- Macro grabada para cálculo del precio por unidad (`PRECIO/UD`)
- Formato condicional sobre ventas (coloreado rojo/amarillo según el monto)

## 💡 Cómo usar

1. Abre `informe-ventas-excel.xlsx` en Microsoft Excel
2. Ve a la hoja `C.mando` y selecciona el año en la celda `B3` para ver los resultados actualizados
3. Explora las hojas:
   - `Subtotales por Región`
   - `Subtotales por Vendedor`
   - `Quitar duplicados`
4. Revisa el gráfico y las tablas dinámicas en la hoja `Segunda Parte`

## 📷 Vista previa

![Vista previa del panel](Tablas_Dinámicas_Gráfico.pdf)

## 🧑‍💻 Autor

- Nombre: Maria Jose Suarez
- Contacto: https://www.linkedin.com/in/suarezmariajose/

---

Este proyecto demuestra cómo Excel puede ser una herramienta poderosa para el análisis de datos sin necesidad de software adicional como Power BI.
