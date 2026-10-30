# Original User Request

## 2026-09-26T18:55:15Z

Auditoría integral profunda de la lógica de negocio, arquitectura y código fuente del sistema de gestión y control de exámenes médicos laborales (Google Apps Script / HTML / JS). Incluye la redacción exhaustiva del ciclo de vida del examen ("Vida del Examen") con diagramas Mermaid, auditoría técnica de cuotas y rendimiento, y una matriz priorizada de mejoras recomendadas. Ninguna modificación de código debe ser implementada en los archivos existentes del proyecto hasta contar con el OK explícito del usuario.

Working directory: c:\Users\cript\OneDrive\Escritorio\Antigravity\Antigravity.Auditoria 2026
Integrity mode: development

## Requirements

### R1. Auditoría de la Lógica de Negocio y Coherencia Operativa
- Realizar un análisis exhaustivo y forense de las reglas de negocio en `Data.js`, `Sync.js`, `Código.js`, `AppLogic.html`, `Desvios.html`, `Avance.html`, `Cobertura.html`, `Reclamos.html`, `Comunicacion.html` y `Reportes.html`.
- Mapear y validar la coherencia de estados de exámenes, transiciones válidas/inválidas, cálculos métricos (porcentajes de avance, cobertura), detección de desvíos, gestión de reclamos y sincronización bidireccional.
- Identificar casos de borde, inconsistencias entre lógica de servidor (GAS) y cliente (JS/HTML), y supuestos no controlados en los datos.

### R2. Redacción Técnica: "Vida del Examen" (Ciclo de Vida Completo)
- Documentar en un archivo Markdown dedicado (`VIDA_DEL_EXAMEN.md`) el ciclo de vida integral de un examen médico de punta a punta.
- Detallar cada fase:
  1. Origen e ingesta de datos del empleado y asignación del examen.
  2. Planificación, agendamiento y estados iniciales.
  3. Ejecución, carga de resultados y transiciones de estado.
  4. Disparadores de desvíos (fechas vencidas, faltantes de documentación, observaciones).
  5. Circuito de reclamos y reasignaciones.
  6. Notificaciones y comunicaciones automáticas/manuales.
  7. Cierre, archivo, consolidación de métricas y reportería.
- Incluir diagramas de flujo y secuencia en sintaxis Mermaid (`flowchart` / `sequenceDiagram` / `stateDiagram-v2`).

### R3. Auditoría Técnica de Código y Rendimiento
- Evaluar el cumplimiento de buenas prácticas en Google Apps Script: límites de ejecución (6 min), cuotas de llamadas a Google Sheets (`SpreadsheetApp.flush()`, lecturas/escrituras en lote vs celda por celda), manejo de caché (`CacheService`/`PropertiesService`) y concurrencia.
- Auditar la seguridad del frontend/backend: validación de inputs, sanitización de datos, manejo seguro de sesiones y errores en `google.script.run`.
- Identificar duplicación de código, acoplamiento excesivo y deuda técnica en scripts de cliente y servidor.

### R4. Catálogo y Matriz de Mejoras Propuestas (Modo Solo Lectura)
- Elaborar un informe estructurado (`PROPUESTAS_DE_MEJORA.md`) con recomendaciones concretas clasificadas por:
  - Categoría: Rendimiento/GAS, Arquitectura & Modularización, Experiencia de Usuario (UI/UX), Robustez/Validaciones.
  - Prioridad e Impacto vs. Esfuerzo (Matriz de prioridad: Alta/Media/Baja).
- Cada propuesta debe incluir: problema identificado, solución técnica sugerida, archivos involucrados y riesgos asociados.
- **DIRECTIVA CRÍTICA**: No modificar ni sobreescribir ningún archivo de código fuente del proyecto (`*.js`, `*.html`, `*.json`) en esta fase. Todos los entregables deben ser informes y documentación.

## Acceptance Criteria

### Entregables de Documentación y Auditoría
- [ ] Archivo `VIDA_DEL_EXAMEN.md` generado con el recorrido cronológico, estados, disparadores y diagramas Mermaid completos y renderizables.
- [ ] Archivo `AUDITORIA_LOGICA_Y_CODIGO.md` generado con el reporte detallado de hallazgos en backend GAS y módulos frontend, clasificados por severidad.
- [ ] Archivo `PROPUESTAS_DE_MEJORA.md` generado con la matriz de mejoras priorizadas, listas para decisión del usuario.
- [ ] Verificación de no modificación: Todos los archivos de código fuente originales (`Data.js`, `Sync.js`, `Código.js`, `*.html`) se mantienen intactos.
