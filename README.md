# BETA INTEGRADO

Suite de sistemas y herramientas internas para apoyar procesos operativos y administrativos de **IGP METALES**.

| Información | Detalle |
| Creador | **Ing. Isidro Vera** |
| Organización autorizada | **IGP METALES** |
| Repositorio “beta-integrado” |
| Uso | Interno y sujeto a autorización de IGP METALES |

## Aviso de uso

El uso de este software está autorizado para **IGP METALES**. Su acceso, copia, modificación o distribución fuera de la organización requiere autorización expresa de IGP METALES y del creador. Este aviso no concede permisos de uso a terceros.

El repositorio contiene código fuente y recursos de los módulos. Las bases de datos operativas, credenciales, llaves privadas, documentos personales, registros y respaldos deben mantenerse fuera de Git y almacenarse en ubicaciones autorizadas y protegidas.

## Módulos

El espacio de trabajo reúne los siguientes proyectos y herramientas:

- **Control de combustible**: aplicación web para operaciones de combustible. Incluye flujos de compras, despachos, consultas y báscula.
- **Control de Calidad *: módulo de control y consulta de calidad.
- **Sistema de depósito IGP METALES**: herramientas asociadas a procesos de depósito.
- **Sistema de Evaluación de Colaboradores**: herramientas de evaluación de colaboradores.
- **Generador de plantillas**: utilidades para generación de plantillas.
La disponibilidad y configuración de cada módulo dependen de su instalación en el equipo correspondiente. Los requisitos detallados a continuación describen **Control de combustible**, según la documentación incluida en `requisitos.txt`; los demás módulos pueden tener requisitos adicionales.

## Requisitos del servidor

Configuración recomendada para el equipo que ejecuta **Control de combustible**:

- Windows 10 u 11 de 64 bits.
- Python 3.14.4.
- Procesador Intel Core i3 o equivalente.
- 4 GB de RAM como mínimo; 8 GB recomendados.
- 10 GB de espacio libre como mínimo, considerando imágenes, documentos y respaldos locales.
- Conexión de red estable y puerto local `5000` disponible.
- Tailscale instalado, conectado y configurado para la red autorizada.
- Permisos para ejecutar PowerShell y mantener activo el servicio.

### Dependencias principales

Las versiones registradas en los requisitos de **Control de combustible** incluyen:

- Flask 3.0.3
- Flask-Login 0.6.3
- Flask-SQLAlchemy 3.1.1
- SQLAlchemy 2.0.51
- ReportLab 4.2.0
- openpyxl 3.1.5
- xlrd 2.0.2
- Pillow 12.3.0
- qrcode

Consulta “requisitos.txt” para el detalle de instalación y compatibilidad.

## Requisitos de los equipos cliente

Para acceder mediante navegador:

- Windows, Linux, macOS o ChromeOS con navegador actualizado.
- Google Chrome, Microsoft Edge, Firefox o Safari (según el dispositivo).
- JavaScript y cookies habilitados.
- 2 GB de RAM como mínimo; 4 GB recomendados.
- Resolución recomendada de 1366 × 768 o superior.
- Conexión a la red autorizada o acceso por Tailscale.

Los formularios y consultas principales pueden utilizarse desde teléfonos y tabletas. Para reportes extensos, tablas grandes y tareas administrativas se recomienda una computadora o tableta.

## Báscula USB/Serial

La lectura desde una báscula requiere un dispositivo compatible con Web Serial y Google Chrome o Microsoft Edge actualizado. Se recomienda una PC o notebook con conexión USB o adaptador USB-Serial. La configuración indicada para la báscula es normalmente **9600 baudios**. El navegador solicitará permiso para acceder al puerto.

En dispositivos o navegadores sin soporte Web Serial, el peso puede ingresarse manualmente si el flujo de trabajo lo permite.

## Red y acceso

Para acceder al sistema de combustible, el servidor debe estar encendido, la aplicación en ejecución y Tailscale conectado. Solo deben acceder usuarios pertenecientes a la red autorizada.

La dirección interna documentada actualmente para **Control de combustible** es:

`http://100.90.6.40:5000`

La dirección puede cambiar según la configuración de red. No es una dirección pública; no habilites acceso externo sin autorización de IGP METALES.

## Datos y respaldos

La base SQLite de **Control de combustible** se almacena localmente en:

`Control de combustible/instance/combustible.db`

Los datos operativos, documentos y respaldos no se incluyen en este repositorio. Mantén copias periódicas en un almacenamiento aprobado por IGP METALES, con acceso restringido y protección adecuada. Verifica las copias según el procedimiento interno de la organización.
