# Convenciones de código — QC-LIMS

## 1. Nombres
- Módulos: `modNombre`
- Clases: `CNombre`
- Formularios: `frmNombre`
- Variables locales: camelCase
- Constantes: MAYÚSCULAS_CON_GUIONES

## 2. Estructura
- Option Explicit obligatorio.
- Una responsabilidad por módulo.
- Formularios sin lógica de negocio.

## 3. Comentarios
- Explicar el "por qué", no el "qué".
- Encabezado obligatorio por módulo.

## 4. Errores
- Usar manejo controlado.  
- Nada de `On Error Resume Next` excepto en validaciones internas.

## 5. Estrategia de pruebas
- Rubberduck Test Modules para cada módulo principal.

