# Arquitectura del proyecto QC-LIMS

QC-LIMS implementa una arquitectura modular sobre Excel + VBA, con separación clara entre:

- Parsing de actividades
- Determinación de analista
- Generación y ciclo de vida de OT
- Validaciones
- Logs y auditoría
- Maestros dinámicos

---

## 1. Visión general

El sistema opera sobre una planilla semanal, donde cada celda puede contener texto libre representando una actividad analítica.

La arquitectura se compone de:

### ✔ Capa de extracciones (`modExtractorIndustrial`)
Detecta actividades en celdas, normaliza textos y genera instancias `CActividadOT`.

### ✔ Capa de parsing (`modParserIndustrial`)
Interpreta texto y lo divide en campos:
- Ensayo
- Técnica
- Muestra
- Lote
- Especialidad

### ✔ Capa de analistas (`modAnalistas`)
Resuelve el analista según bloque de filas de la planilla semanal.

### ✔ Capa de OT (`modOT`)
Genera Órdenes de Trabajo con identificador único y registra sus estados.

### ✔ Capa de logs
Registra toda acción crítica del sistema.

### ✔ Formularios (`forms/`)
Interacción con usuario, selección de actividades, filtros y confirmaciones.

---

## 2. Diagrama conceptual

Celda → Parser → Actividad (CActividadOT) → Analista → OT → Estado → Log

---

## 3. Objetivos de arquitectura

- Separación estricta de responsabilidades (SRP).
- Evitar lógica mezclada en formularios.
- Código testeable con Rubberduck.
- Preparación futura para migración a .NET o web.

---

## 4. Dependencias

- Excel/VBA
- Rubberduck (inspecciones, refactor, pruebas unitarias)
- Git para versionado


