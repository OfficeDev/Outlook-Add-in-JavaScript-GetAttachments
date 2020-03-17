---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: En este ejemplo, se muestra cómo obtener datos adjuntos de un buzón de Exchange.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Complemento de Outlook: Obtener datos adjuntos desde un servidor Exchange

**Tabla de contenido**

* [Resumen](#summary)
* [Requisitos previos](#prerequisites)
* [Componentes clave del ejemplo](#components)
* [Descripción del código](#codedescription)
* [Compilar y depurar](#build)
* [Solución de problemas](#troubleshooting)
* [Preguntas y comentarios](#questions)
* [Recursos adicionales](#additional-resources)

<a name="summary"></a>
##Resumen
En este ejemplo, se muestra cómo obtener datos adjuntos de un buzón de Exchange.

<a name="prerequisites"></a>
## Requisitos previos ##

Este ejemplo necesita lo siguiente:  

  - Visual Studio 2013 con Update 5 o Visual Studio 2015.  
  - Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de Office 365. Puede [participar en el programa para desarrolladores Office 365 y obtener una suscripción gratuita durante 1 año a Office 365](https://aka.ms/devprogramsignup).
  - Cualquier explorador que admita ECMAScript 5.1, HTML5 y CSS3, como Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 o una versión posterior de estos exploradores.
  - Familiaridad con los servicios web y la programación de JavaScript.

<a name="components"></a>
## Componentes clave del ejemplo
La solución de ejemplo contiene los archivos siguientes:

- AttachmentExampleManifest.xml: El archivo de manifiesto del complemento de Outlook.
- AppRead\Home\Home.html: La interfaz de usuario HTML para el complemento de correo electrónico de Outlook.
- AppRead\Home\Home.js: El archivo JavaScript que procesa el envío de la información de datos adjuntos al servicio de datos adjuntos remoto que se incluye con este ejemplo.

El proyecto AttachmentService define un servicio REST mediante el uso de la API de WCF. El proyecto contiene los siguientes archivos:

- Controllers\AttachmentServiceController.cs: El objeto de servicio que proporciona la lógica empresarial del servicio de ejemplo.
- Models\ServiceRequest: El objeto que representa una solicitud web. El contenido del objeto se crea desde un objeto de solicitud JSON enviado desde el complemento de correo electrónico.
- Models\Attachment.cs: El objeto de utilidad que ayuda a deserializar el objeto JSON que envía el complemento de correo electrónico.
- Models\AttachmentDetails.cs: El objeto que representa los detalles de cada dato adjunto. Proporciona un objeto .NET Framework que coincide con el objeto `AttachmentDetails` del complemento de correo electrónico.
- Models\ServiceResponse: El objeto que representa una respuesta del servicio web. El contenido del objeto se serializa en un objeto JSON cuando se vuelve a enviar al complemento de correo electrónico.
- Web.config: Enlaza el servicio de ejemplo al extremo de servidor web.



<a name="codedescription"></a>
##Descripción del código

Este ejemplo le muestra cómo recuperar los datos adjuntos de un servicio web que sea compatible con el complemento de correo electrónico. Por ejemplo, puede crear un servicio que cargue fotos en un sitio de uso compartido o un servicio que almacene documentos en un repositorio. El servicio obtiene los datos adjuntos directamente desde el servidor de Exchange y no requiere que el cliente realice ningún proceso adicional para obtener los datos adjuntos y luego enviarlos al servicio.

Este ejemplo tiene dos partes. La primera, la aplicación de correo, se ejecuta en el cliente de correo electrónico. El complemento de correo se mostrará siempre que el elemento activo sea un mensaje o una cita. Cuando selecciona el botón **Probar datos adjuntos**, el complemento de correo envía detalles sobre el archivo adjunto al servicio web que procesa la solicitud. El servicio emplea los siguientes pasos para procesar los datos adjuntos:

- Envía una solicitud de operación [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) al servidor de Exchange que hospeda el buzón. El servidor responde enviando los datos adjuntos al servicio. En este ejemplo, el servicio simplemente escribe el XML desde el servidor hacia los resultados de seguimiento.
- Devuelve el número de datos adjuntos procesados a la aplicación de correo.



<a name="build"></a>
## Compilar y depurar ##
**Nota:** El complemento de correo se activará en cualquier mensaje de correo electrónico de la bandeja de entrada del usuario que tenga uno o más datos adjuntos. Puede hacer que sea más fácil probar el complemento enviando uno o más mensajes de correo electrónico a la cuenta de prueba antes de ejecutar el ejemplo de complemento.

1. Abra la solución en Visual Studio.
2. Haga clic derecho sobre la solución en el Explorador de soluciones. Seleccione **establecer proyectos de inicio**. 
3. Seleccione **Propiedades comunes**, y elija **Proyecto de inicio**.
4. Asegúrese de que la **Acción** para el proyecto **AttachmentExampleService** esté establecida en **Iniciar**.
5. Pulse F5 para crear e implementar el complemento de ejemplo.
6. Conecte a una cuenta de Exchange proporcionando la dirección de correo electrónico y la contraseña de un servidor de Exchange 2013.
7. Permita que el servidor configure la cuenta de correo.
8. Inicie sesión en la cuenta de correo electrónico escribiendo el nombre de la cuenta y la contraseña. 
9. Selecciones un mensaje en la Bandeja de entrada
10. Espere a que aparezca la barra del complemento sobre el mensaje.
11. En la barra del complemento, haga clic en **AttachmentExample**.
12. Cuando aparezca el complemento de correo, haga clic en el botón **TestAttachments** para enviar una solicitud al servidor de Exchange.
13. El servidor responderá con el número de datos adjuntos procesados para el elemento. Esto debería ser igual al número de datos adjuntos que contiene el elemento.

<a name="troubleshooting"></a>
##Solución
de problemas A continuación se enumeran los errores comunes que pueden ocurrir al usar Outlook Web App para probar un complemento de correo para Outlook:

- La barra de complemento no aparece cuando se selecciona un mensaje. Si esto ocurre, vuelva a iniciar la aplicación seleccionando **Depuración: detener depuración** en la ventana de Visual Studio y presione F5 para recompilar e implementar el complemento. 
- Es posible que los cambios en el código de JavaScript no se hayan recogido al implementar y ejecutar el complemento. Si no se han añadido los cambios, borre la memoria caché en el explorador web. Para ello, seleccione **Herramientas: opciones de Internet** y haga clic en el botón **Eliminar...**. Elimine los archivos temporales de Internet y reinicie el complemento. 

<a name="questions"></a>
##Preguntas y comentarios##

- Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues).
- Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].


<a name="additional-resources"></a>
## Recursos adicionales ##

- [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [API web: El sitio oficial de Microsoft ASP.NET](http://www.asp.net/web-api)
- [Procedimiento: Obtener datos adjuntos desde un servidor Exchange](http://msdn.microsoft.com/library/dn148008.aspx)

## Derechos de autor
Copyright (c) 2015 Microsoft. Todos los derechos reservados.


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
