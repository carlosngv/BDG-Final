# Verificación de Licencias

La verificación de licencias dentro de la empresa se vuelve una tarea supremamente repetitiva y que a simple vista se puede decir que puede ser automatizada. De cualquier forma, en este caso utilizando la herramienta UIPath.

En el presente programa, se llevo a cabo una tarea que, mediante un archivo Excel con todas las licencias de los clientes y demás información, se verifica si la licencia ha caducado o no.




## Función del programa

El programa inicialmente se redirige al archivo Excel compartido, realiza la descarga y este es guardado con la extensión ".xlsx".


El programa abre el archivo en el directio de descargas del usuario. Itera cada fila del archivo y verifica las licencias por cada cliente, si este está a 5 meses de vencer o ya estén vencidads, serán notificadas al cliente mediante un agradable correo, que lleva a cabo está función.

Dentro del archivo de licencias hay dos hojas: Licencias e Historial. En licencias se hace la validación, las licencias vencidas, menores a 5 meses son llevadas al historial y se notifican.
