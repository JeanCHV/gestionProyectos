\# Requerimientos del Sistema de Gestión de Riesgos



\## Requerimientos Funcionales



\### Gestión de Activos



\- \*\*RF-01:\*\* El sistema debe permitir registrar activos manualmente.

\- \*\*RF-02:\*\* El sistema debe permitir importar activos desde un archivo Excel.

\- \*\*RF-03:\*\* El sistema debe validar que el archivo Excel contenga las columnas obligatorias: ID, nombre del activo, descripción, tipo, responsable y valor del activo.

\- \*\*RF-04:\*\* El sistema debe permitir consultar, editar y eliminar activos registrados.

\- \*\*RF-05:\*\* El sistema debe evitar el registro de activos duplicados.



\### Identificación de Riesgos



\- \*\*RF-06:\*\* El sistema debe permitir registrar riesgos asociados a un activo.

\- \*\*RF-07:\*\* Cada riesgo debe incluir ID, nombre, descripción, causa, consecuencia, activo afectado y responsable.

\- \*\*RF-08:\*\* El sistema debe permitir clasificar el riesgo según su origen (técnico, operativo, financiero, legal, seguridad, recurso humano, proveedor, entre otros).

\- \*\*RF-09:\*\* El sistema debe permitir definir el horizonte temporal del riesgo (corto, mediano o largo plazo).



\### Valoración de Riesgos



\- \*\*RF-10:\*\* El sistema debe permitir asignar valores de probabilidad e impacto para cada riesgo.

\- \*\*RF-11:\*\* El sistema debe calcular automáticamente el nivel de riesgo utilizando la fórmula:



&#x20; ```

&#x20; Riesgo = Probabilidad × Impacto

&#x20; ```



\- \*\*RF-12:\*\* El sistema debe clasificar automáticamente el riesgo como Bajo, Medio o Alto.

\- \*\*RF-13:\*\* El sistema debe permitir configurar los rangos utilizados para la clasificación del nivel de riesgo.



\### Matriz de Riesgos



\- \*\*RF-14:\*\* El sistema debe generar una matriz de calor de riesgos.

\- \*\*RF-15:\*\* El sistema debe ubicar automáticamente cada riesgo según su probabilidad e impacto.

\- \*\*RF-16:\*\* Cada riesgo debe visualizarse mediante su identificador dentro de la matriz.

\- \*\*RF-17:\*\* La matriz debe diferenciar visualmente los niveles Bajo, Medio y Alto.

\- \*\*RF-18:\*\* El sistema debe permitir filtrar la matriz por activo, tipo de riesgo, responsable, plazo y nivel de criticidad.



\### Plan de Mitigación



\- \*\*RF-19:\*\* El sistema debe permitir registrar un plan de mitigación para cada riesgo.

\- \*\*RF-20:\*\* El plan de mitigación debe incluir acción preventiva, acción correctiva, responsable, fechas de inicio y fin, recursos necesarios y estado.

\- \*\*RF-21:\*\* El sistema debe permitir definir la estrategia de tratamiento del riesgo (Evitar, Mitigar, Transferir o Aceptar).

\- \*\*RF-22:\*\* El sistema debe permitir realizar seguimiento al avance del plan de mitigación.

\- \*\*RF-23:\*\* El sistema debe permitir adjuntar evidencias del cumplimiento del plan.



\### Priorización de Riesgos



\- \*\*RF-24:\*\* El sistema debe ordenar los riesgos según su nivel de criticidad.

\- \*\*RF-25:\*\* El sistema debe mostrar una lista priorizada de riesgos.

\- \*\*RF-26:\*\* El sistema debe permitir marcar riesgos como críticos o urgentes.

\- \*\*RF-27:\*\* El sistema debe recomendar el orden de atención considerando probabilidad, impacto y plazo.



\### Seguimiento y Control



\- \*\*RF-28:\*\* El sistema debe permitir actualizar el estado del riesgo (Identificado, Evaluado, En tratamiento, Mitigado, Aceptado o Cerrado).

\- \*\*RF-29:\*\* El sistema debe permitir registrar revisiones periódicas del riesgo.

\- \*\*RF-30:\*\* El sistema debe recalcular automáticamente el riesgo después de aplicar controles.

\- \*\*RF-31:\*\* El sistema debe diferenciar entre riesgo inherente y riesgo residual.

\- \*\*RF-32:\*\* El sistema debe generar alertas cuando un riesgo de nivel Alto no tenga un plan de mitigación.



\### Reportes



\- \*\*RF-33:\*\* El sistema debe generar reportes de riesgos por proyecto.

\- \*\*RF-34:\*\* El sistema debe permitir exportar reportes en formato Excel y PDF.

\- \*\*RF-35:\*\* El sistema debe mostrar indicadores de riesgos clasificados por nivel.

\- \*\*RF-36:\*\* El sistema debe mostrar el avance de los planes de mitigación.

\- \*\*RF-37:\*\* El sistema debe generar un resumen ejecutivo de los riesgos críticos.



\---



\## Requerimientos No Funcionales



\- \*\*RNF-01:\*\* El sistema debe contar con una interfaz intuitiva y fácil de usar.

\- \*\*RNF-02:\*\* El sistema debe validar la integridad y consistencia de la información ingresada.

\- \*\*RNF-03:\*\* El sistema debe gestionar el acceso mediante roles (Administrador, Gestor de Riesgos, Responsable del Activo y Auditor).

\- \*\*RNF-04:\*\* El sistema debe proteger la información mediante autenticación y autorización.

\- \*\*RNF-05:\*\* El sistema debe registrar un historial de cambios para garantizar la trazabilidad.

\- \*\*RNF-06:\*\* El sistema debe responder en tiempos adecuados durante la consulta de la matriz de riesgos y la generación de reportes.

\- \*\*RNF-07:\*\* El sistema debe permitir realizar copias de seguridad de la información.

\- \*\*RNF-08:\*\* El sistema debe ser escalable para soportar múltiples proyectos y grandes volúmenes de riesgos.

