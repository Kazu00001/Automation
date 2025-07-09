# Automation
1-. Descargar python https://www.python.org/downloads/

2-. Descargar un complilador, yo recomendaria pycharm https://www.python.org/downloads/

3-. Descargar repositorio atravez de un zip en repositorio en el apartado llamado **<Code>**, en la parte inferior 
selecionar descargar zip.

4-. Abrir el repositorio con el complilardor

5-. Descargar Librerias en la secion de importacion del archivo Automation en las primeras lineas
que comienzan con un import presionar las teclas (**Presiona Alt + Enter** Para windows) y (**Command + .**, Para dispositivos con ios)
/ otra manera es ejecutando el siguiente comando en terminal verificar que pip este instalado primero: **pip --version** y despues: **pip install polars python-docx**




Cosas importantes a comentar:
1-. El codigo funciona apartir del exel para su correcto funcionamiento eliminar el contenido de Save_Files y eliminar el archivo de
empleados el codigo se encarga de crearlo.
2-. el archivo exel debe de tener el mismo nombre o si se cambia el nombre hacer la modificacion en la linea #15
3-. Otro punto a considerar es que el programa le el excel base a encabezados especificos:
Nombre del Empleado,	CURP,	Ocupacion	Puesto,	Nombre de la Empresa,	Registro SHCP (RFC),	Nombre del Programa,	Duración del Programa, Fecha de Inicio,	Fecha de Fin,	Area Temática del Curso,	Nombre del Agente Capacitador Asegurarce de tener la infomacion en la correcta celda y verificar el nombre de la celda.
