# Automatizacion-
"Automatización del flujo de Control de Calidad (QA). Conecta múltiples bases de datos de Google Sheets para generar reportes dinámicos y feedback inmediato a agentes mediante correos electrónicos personalizados."

📊 QA Automation: Clinicals Call Verification
Este repositorio contiene la lógica de automatización para el proceso de Control de Calidad (QA) de Clinicals. El sistema integra Google Sheets con un motor de notificaciones dinámicas para optimizar el seguimiento de auditorías de llamadas.

🎯 Objetivo del Proyecto
Digitalizar y automatizar el flujo de trabajo de auditoría, eliminando la entrada manual de datos y asegurando que los agentes reciban retroalimentación inmediata sobre su desempeño.

🚀 Funcionalidades Principales
Sincronización Multifuente: Localiza registros automáticamente en bases de datos externas (Scorecard Data, No records) utilizando el identificador MRN.

Motor de Notificaciones Dinámicas:

Destinatario Inteligente: Envía el reporte automáticamente al correo extraído del campo Agent Email.

Copia de Supervisión: Incluye en CC al supervisor correspondiente y al equipo de administración.

Asunto Personalizado: Genera líneas de asunto únicas basadas en el nombre del agente auditado.

Visualización de Datos (HTML Email):

Alertas Visuales: Resaltado automático en rojo para resultados "Not Valid" o "Red Flag: YES".

Gestión de Multimedia: Genera enlaces directos a las grabaciones de audio (Agent Recording) almacenadas en Google Drive.

Interfaz Profesional: Diseño responsivo con logotipos corporativos integrados.

🛠️ Stack Tecnológico
Lenguaje: JavaScript / Google Apps Script.

Integraciones: Google Sheets API, Gmail App Service, Google Drive API.

Frontend: HTML5 / CSS3 para el modal de captura de datos y plantillas de correo.

📂 Estructura del Repositorio
Codigo.js: Contiene el núcleo de la lógica (Procesamiento de formularios, búsqueda de MRN y envío de correos).

Frontend/: Directorio con la interfaz de usuario para los auditores.
