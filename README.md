# 📦 Stock Dashboard — Tarifa Nacional vs Stock Global

Dashboard Streamlit para análisis de stock en almacén cruzado con la tarifa nacional.

## Instalación

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fuentes de datos

La app acepta **dos modos** de carga:

### Modo 1: Subir archivos
Sube directamente desde tu ordenador:
- `TARIFA_NACIONAL_COMPLETA.xlsx` — con hojas `T_AMZ`, `T_MIR`, `T_C4`, `T_MM`, `T_PRIV`
- `Stock_Global.xlsx` — con hoja `España` (y opcionalmente `Alemania`, `Francia`, `Italia`)

### Modo 2: URLs remotas (recomendado para producción)
Introduce las URLs de descarga directa en el sidebar. Las URLs se introducen en tiempo real
y **nunca se guardan en el código**, lo que permite que la app sea pública sin exponer rutas.

> Las URLs deben ser de descarga directa (no Google Drive preview). 
> Para Google Drive: botón derecho → "Copiar enlace de descarga directa"

## Estructura esperada

### Tarifa (`T_AMZ`, `T_MIR`, etc.)
| REFERENCIA | EAN | NOMBRE COMPLETO | FAMILIA | SUBFAMILIA | PVPR | NETO | ... |

### Stock España
| Referencia | Descripcion | Familia Padre | Familia | Subfamilia | Stock Fisico | Stock Operativo | Stock Comercial | Mar | Puerto | Despachado |

## Vistas del dashboard

| Tab | Descripción |
|-----|-------------|
| 📊 Vista General | Tabla resumen por familia + donut de estados |
| 🔴 Sin Stock | SKUs sin stock: riesgo real vs. con entrante |
| 🚢 Tránsito | Stock en mar, en puerto, despachado |
| 📈 Por Categoría | Gráficos de barras por familia/subfamilia |
| 🔎 Búsqueda SKU | Buscador por referencia o nombre |

## Semáforo de estados

- 🔴 **Sin stock** — Stock Comercial = 0 y sin ningún entrante (rotura inminente)
- 🟡 **Sin stock (entrante)** — Stock = 0 pero hay mercancía en mar/puerto/despachado
- 🟠 **Stock bajo** — Stock Comercial entre 1 y 3 unidades
- 🟢 **OK** — Stock suficiente

## Canales disponibles

- Amazon (AMZ)
- Miravia (MIR)  
- C4 / Corte Inglés (C4)
- MediaMarkt (MM)
- Privalia (PRIV)
