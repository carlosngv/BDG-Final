- [Verificación de Licencias](#Verificación-de-Licencias)
- [Función del programa](#Función-del-programa)
- [Envio de correo](#Envio-de-correo)

## Verificación de Licencias

La verificación de licencias dentro de la empresa se vuelve una tarea supremamente repetitiva y que a simple vista se puede decir que puede ser automatizada. De cualquier forma, en este caso utilizando la herramienta UIPath.

En el presente programa, se llevo a cabo una tarea que, mediante un archivo Excel con todas las licencias de los clientes y demás información, se verifica si la licencia ha caducado o no.




## Función del programa

El programa inicialmente se redirige al archivo Excel compartido, realiza la descarga y este es guardado con la extensión ".xlsx".


El programa abre el archivo en el directio de descargas del usuario. Itera cada fila del archivo y verifica las licencias por cada cliente, si este está a 5 meses de vencer o ya estén vencidads, serán notificadas al cliente mediante un agradable correo, que lleva a cabo está función.

Dentro del archivo de licencias hay dos hojas: Licencias e Historial. En licencias se hace la validación, las licencias vencidas, menores a 5 meses son llevadas al historial y se notifican.

## Envio de correo

El programa genera un cuerpo de html en el que se almacena en la variable cuerpoCorreo con la siguiente estructura, adjuntado los datos de Codigo de Licencia, Proveedor, Tipo de Licencia, Fecha de Inicio y Fecha de Fin.

El correo se envia mediante la actividad Send Email, Outlook, en el cual se agrega el cuerpo del correo como html.

    "<html>" + _
    "<body>" + _
    "<H1  style=' color : #1f497d;'>NOTIFICACIÓN DE LICENCIA BDG, S.A.</H1>" + _
    "<table Class='Default' style=' color : #1f497d;' border=2>" + _
    "<tr>" + _
    "<td>Codigo de Liecencia</td>" + _
    "<td>Proveedor</td>" + _
    "<td>Tipo de Licencia</td>" + _
    "<td>Fecha de Inicio</td>" + _
    "<td>Fecha de Fin</td>" + _
    "</tr>" + _
    "<tr>" + _
    String.Format( 
    "<td>{0}</td>" + _
    "<td>{1}</td>" + _
    "<td>{2}</td>" + _
    "<td>{3}</td>" + _
    "<td>{4}</td>" + _
    "</tr>" + _
    "</table>", row.Item("Codigo de licencia"), proveedorStr, row.Item("Tipo de licencia"), proveedorStr, row.Item("Fecha Inicio"), proveedorStr, row.Item("Fecha Fin")) + _
    "<br>" + _
    "<div id='nombre' style=' color : #B91D1D;'>" + _
    "Anthony Son &" + _
    "Carlos Ojani" + _
    "</div>" + _
    "<div id='puesto' style=' color : #ADADAD;'>" + _
    "Programa Trainee" + _
    "</div>" + _
    "<img src='https://export.com.gt/attach/componentes/slider-infinito/slider-infinito-5-5c192fafb3128.jpg' style='width:2.625in;height:.6354in'>" + _
    "</body>" + _
    "</html>"