# QC-LIMS  
Sistema modular de gestión de actividades analíticas para Laboratorios de Control de Calidad (LCC)

QC-LIMS es un LIMS artesanal desarrollado sobre Excel + VBA, diseñado específicamente para:
- gestionar actividades analíticas semanales,
- asignar cargas de trabajo a analistas,
- generar Órdenes de Trabajo (OT),
- registrar el ciclo de vida de cada ensayo,
- integrar maestros dinámicos y lógica de parsing industrial.

El proyecto evoluciona hacia una arquitectura modular, versionada mediante Git y preparada para futura migración a un backend real.

---

## 🚀 Objetivos principales

- Centralizar y estructurar la planificación semanal del laboratorio.  
- Estandarizar la carga, asignación y seguimiento de actividades.  
- Mantener un registro completo de decisiones y estados.  
- Facilitar la trazabilidad y la auditoría.  
- Servir como base para una futura digitalización completa del LCC.

---

## 🧩 Componentes del sistema

- **Parser Industrial**  
  Extrae ensayos, técnicas, muestras, lotes y especialidades desde texto libre en celdas Excel.

- **Gestor de Analistas**  
  Determina el analista responsable según bloque de planilla.

- **Generador de Órdenes de Trabajo (OT)**  
  Agrupa actividades seleccionadas, asigna número único y registra el ciclo de vida.

- **Log del Sistema**  
  Cada acción relevante deja registro permanente.

---

## 📁 Estructura del repositorio

QC-LIMS/
│
├─ src/
│ ├─ modules/ ' Módulos .bas
│ ├─ classes/ ' Clases .cls
│ └─ forms/ ' Formularios .frm + .frx
│
├─ docs/
│ ├─ arquitectura.md
│ ├─ roadmap.md
│ ├─ decisiones.md
│ └─ convenciones_codigo.md
│
└─ README.md


---

## 🔧 Requisitos

- Excel + VBA
- Rubberduck 2.5+
- Git (opcional pero recomendado)
- Windows

---

## 🧪 Estado actual del proyecto

- Parser industrial → ✔ estable  
- Gestión de analistas → ✔ corregida  
- Generador de OT → ✔ operativo  
- Ciclo de estados y reversión de OT → 🔄 en desarrollo  
- Validaciones cruzadas / duplicados → 🔄 planificadas  

---

## 👤 Autor

Proyecto desarrollado por Matías Olivera, junto con asistencia técnica de ChatGPT.  
