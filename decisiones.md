# Decisiones de diseño — QC-LIMS

Este documento registra decisiones arquitectónicas clave tomadas durante el desarrollo.

---

## D1 — Determinar analista por bloque (filaDesde/filaHasta)
Motivo: robustez y simplicidad comparado con leer títulos en planilla.

Estado: ✔ implementado.

---

## D2 — Parser descentralizado y modular
Separar extracción del parsing evita errores circulares y permite test unitarios.

Estado: ✔ estable.

---

## D3 — OT con estados diferenciados
Define ciclo de vida realista para un laboratorio:
- pendiente
- en_proceso
- finalizada
- anulada
- cancelada

Estado: 🟡 en implementación.

---

## D4 — Reversión controlada de actividades
Una OT anulada devuelve actividades a estado “libre”.  
Una OT cancelada deja actividades inutilizables.

---

## D5 — Identificación única de actividad
Combinación:

Especialidad + Ensayo + Técnica + Lote


Evita duplicaciones silenciosas.

---

## D6 — Todo debe quedar registrado
Toda acción crítica se escribe en LOG_OT.

