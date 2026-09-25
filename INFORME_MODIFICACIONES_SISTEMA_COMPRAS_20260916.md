# Informe de modificaciones del sistema de compras

**Fecha:** 16 de septiembre de 2026  
**Sistema:** Control de combustible - módulo de Compras  
**Alcance:** Ajustes realizados a partir de los requerimientos planteados en el correo de análisis del proceso de compras.

## 1. Objetivo

Alinear el módulo de Compras con el flujo operativo solicitado por la empresa, manteniendo las funciones que ya estaban disponibles y corrigiendo los puntos que impedían utilizar correctamente la aprobación de Gerencia.

## 2. Funcionalidad que ya existía

Antes de esta intervención, el sistema ya contaba con:

- Roles diferenciados para Compras y Gerencia.
- Redirección de los usuarios de Gerencia a su bandeja de aprobación.
- Registro de pedidos de insumos y servicios.
- Asociación del pedido con el supervisor solicitante.
- Carga de presupuestos y documentos de respaldo.
- Selección de un presupuesto para la aprobación.
- Registro del historial de decisiones de Gerencia.
- Notificaciones a Compras y al solicitante.
- Generación de órdenes de compra luego de una aprobación.
- Seguimiento del estado de las órdenes y servicios.
- Registro de proveedores, formas de pago y clasificación contable.

## 3. Modificaciones realizadas

### 3.1 Corrección del acceso de Gerencia

Se identificó que el acceso fallaba después del inicio de sesión porque la ruta `/gerencia/pedidos` intentaba cargar una plantilla inexistente. Se creó la vista de Gerencia con:

- Bandeja de solicitudes pendientes.
- Detalle del pedido y solicitante.
- Visualización de presupuestos adjuntos.
- Selección del presupuesto aprobado.
- Registro de observaciones.
- Botones de aprobación y rechazo.
- Listado de solicitudes que aún esperan presupuestos.

También se corrigió una expresión de formato incompatible con Jinja que producía un error interno al mostrar las cantidades del pedido.

### 3.2 Corrección de la evaluación del umbral de Gerencia

La solicitud de Compras podía evaluar el umbral con un monto igual a cero, por lo que la decisión no siempre representaba el valor real del pedido.

La lógica fue corregida para evaluar el total estimado real y determinar correctamente cuándo corresponde la intervención de Gerencia.

### 3.3 Excepciones operativas

Se agregó una regla para contemplar situaciones en las que no resulta razonable exigir tres presupuestos:

- Compras urgentes o emergentes.
- Situaciones críticas o de mantenimiento prioritario.
- Casos en los que exista un único proveedor registrado.

Estas excepciones quedan dentro del flujo de aprobación y no eliminan la trazabilidad de la decisión.

### 3.4 Validación y trazabilidad

Se mantuvo el registro de:

- Usuario que aprueba o rechaza.
- Fecha de la decisión.
- Observación de Gerencia.
- Presupuesto seleccionado, cuando corresponde.
- Historial de aprobación asociado al pedido.

## 4. Verificaciones realizadas

Se ejecutaron las siguientes comprobaciones:

- Pruebas automatizadas del flujo de compras: **2 pruebas correctas**.
- Compilación de la aplicación Python: **correcta**.
- Compilación de la plantilla Jinja de Gerencia: **correcta**.
- Acceso autenticado con un usuario de rol `gerencia`: **HTTP 200**.
- Reinicio del servicio activo.
- Respuesta de la aplicación en el puerto `5000`: **operativa**.

## 5. Correspondencia con los requerimientos del correo

| Requerimiento | Resultado actual |
|---|---|
| Participación y aprobación de Gerencia | Implementado y verificado |
| Comparación de presupuestos | Ya existente y operativo |
| Excepción por urgencia | Agregada |
| Excepción por proveedor único | Agregada |
| Trazabilidad de decisiones | Implementada mediante historial y observaciones |
| Notificaciones a los involucrados | Ya existente y conservada |
| Generación de orden de compra | Ya existente y conservada |
| Recepción y seguimiento operativo | Parcialmente existente mediante estados de pedido y orden |
| Gestión avanzada de tipo de cambio | Existe una cotización de referencia; requiere una ampliación específica si se necesita histórico por operación |
| Aprobación por planta o área | El pedido conserva el supervisor solicitante y destinatario; requiere una ampliación adicional si se desea una matriz formal por planta/área |

## 6. Pendientes recomendados y desarrollo propuesto

Los siguientes puntos corresponden a una segunda etapa de desarrollo. No deben interpretarse como funcionalidades ya concluidas; representan la ampliación necesaria para llevar el módulo desde el circuito actual de aprobación y compra hacia un sistema integral de abastecimiento.

### 6.1 Matriz formal de aprobación por planta, área y monto

**Objetivo:** establecer quién puede solicitar, revisar y aprobar una compra según la planta, el área solicitante, el tipo de adquisición y el monto.

**Desarrollo propuesto:**

- Incorporar plantas, áreas y centros de costo como datos del pedido.
- Definir niveles de aprobación configurables por rango de monto.
- Asociar cada área con uno o más responsables autorizados.
- Impedir que una persona apruebe pedidos fuera de su ámbito, salvo usuarios administradores autorizados.
- Registrar cada etapa de aprobación con usuario, fecha, decisión y observación.

**Resultado esperado:** el sistema reflejará la estructura real de la empresa y evitará aprobaciones genéricas que no identifiquen la responsabilidad de cada planta o área.

### 6.2 Registro formal de recepción en Depósito

**Objetivo:** cerrar el circuito entre la orden de compra y la recepción física de los bienes o servicios.

**Desarrollo propuesto:**

- Crear una recepción vinculada a la orden de compra.
- Registrar fecha, responsable, proveedor, número de remisión o documento de respaldo.
- Permitir indicar la cantidad recibida por cada detalle.
- Admitir recepciones parciales cuando el proveedor entregue en más de una fecha.
- Adjuntar remisión, factura, acta de conformidad u otro respaldo.
- Para servicios, registrar la conformidad del área que recibió la prestación.

**Resultado esperado:** Depósito y Compras podrán demostrar qué se recibió, cuándo, en qué cantidad y quién lo validó.

### 6.3 Control de diferencias entre solicitado, comprado y recibido

**Objetivo:** detectar automáticamente diferencias de cantidad, precio, proveedor o alcance.

**Desarrollo propuesto:**

- Comparar el detalle original del pedido con la orden aprobada.
- Comparar la orden de compra con una o varias recepciones.
- Identificar faltantes, sobrantes, entregas parciales y sustituciones.
- Calcular la diferencia entre el monto estimado, el monto aprobado y el monto finalmente recibido o facturado.
- Solicitar observación y autorización cuando exista una variación fuera del margen permitido.
- Mostrar el estado de conciliación: pendiente, parcial, conforme u observado.

**Resultado esperado:** se reducirá el riesgo de cerrar compras con cantidades incorrectas o sin justificar diferencias.

### 6.4 Tipo de cambio histórico por presupuesto y orden de compra

**Objetivo:** conservar el valor de referencia utilizado al momento de cotizar, aprobar y emitir la orden.

**Desarrollo propuesto:**

- Guardar moneda, monto original, tipo de cambio aplicado y monto convertido.
- Registrar la fecha y la fuente del tipo de cambio.
- Mantener el tipo de cambio histórico aunque posteriormente se modifique la configuración general.
- Mostrar los importes en moneda original y en guaraníes cuando corresponda.
- Utilizar el mismo valor histórico en reportes, órdenes y comparaciones de proveedores.

**Resultado esperado:** los reportes conservarán el contexto financiero original y permitirán auditar por qué un presupuesto fue evaluado con determinado valor.

### 6.5 Flujo formal de compras de emergencia

**Objetivo:** permitir una compra urgente sin eliminar los controles ni la justificación posterior.

**Desarrollo propuesto:**

- Incorporar un tipo de solicitud “Emergencia”.
- Exigir motivo, impacto operativo, fecha límite y responsable que solicita.
- Permitir la selección de proveedor único o compra directa con justificación obligatoria.
- Notificar inmediatamente a Gerencia y a Compras.
- Registrar si la autorización fue previa o posterior a la compra.
- Exigir la regularización documental dentro de un plazo configurable.
- Generar un reporte específico de compras de emergencia y sus causas.

**Resultado esperado:** la empresa podrá actuar ante paradas, mantenimientos críticos o faltantes urgentes sin convertir la excepción en un proceso sin trazabilidad.

### 6.6 Reporte consolidado de tiempos, ahorros, proveedores y cumplimiento

**Objetivo:** transformar los datos del circuito de compras en información para la toma de decisiones.

**Desarrollo propuesto:**

- Medir tiempo desde la solicitud hasta la aprobación.
- Medir tiempo desde la aprobación hasta la orden y desde la orden hasta la recepción.
- Comparar presupuesto inicial, presupuesto seleccionado y costo final.
- Calcular ahorro por proveedor, categoría, planta y período.
- Mostrar pedidos demorados, rechazados, urgentes y con diferencias de recepción.
- Evaluar cumplimiento de proveedores respecto de plazo, cantidad y conformidad.
- Permitir filtros por fecha, área, planta, solicitante, proveedor y estado.
- Exportar resultados a Excel o PDF para reuniones de gestión.

**Resultado esperado:** Gerencia contará con indicadores objetivos para evaluar costos, tiempos, proveedores y cuellos de botella del proceso.

### 6.7 Orden sugerido de implementación

Para reducir riesgos y obtener valor progresivamente, se recomienda el siguiente orden:

1. Matriz de aprobación por planta, área y monto.
2. Recepción formal en Depósito.
3. Control de diferencias y conciliación.
4. Tipo de cambio histórico.
5. Flujo formal de emergencias.
6. Reportes e indicadores consolidados.

Este orden prioriza primero la gobernanza y la trazabilidad de la operación; luego incorpora el análisis gerencial sobre la información acumulada.

## 7. Conclusión

El módulo de Compras no partía de una estructura vacía: ya disponía de los elementos principales del circuito de solicitudes, presupuestos, Gerencia y órdenes de compra. Las modificaciones realizadas corrigieron el acceso de Gerencia, estabilizaron la evaluación del monto real y agregaron excepciones operativas para que el flujo responda mejor a situaciones urgentes y de proveedor único.

El sistema queda operativo para continuar utilizando el circuito actual, con una base preparada para incorporar las ampliaciones de control por planta, recepción formal y gestión histórica de moneda en una siguiente etapa.
