# SBOVIDCURRENCY
Aplicación de consola que corre en linux el cual descarga los tipos de cambio al banco central y lo ingresa a un BD MongoDB. Manejar los tipos de cambios del Banco Central con las monedas configuradas en SAPBO conectandose a un Webservice interno. Desarrollado en C#.

La aplicación Tipo de Cambio es una solución externa a SAPB1, la cual se encarga de extraer diariamente los distintos tipos de cambio de la página oficial del Banco Central, con las monedas configuradas en SAPBO conectándose a un Web Servicie interno

La ejecución de esta aplicación se realiza mediante la programación de tareas de Windows, las cuales se programan para ejecutarse a diario después de las 20 horas por lo general, ya que en ese horario el banco actualiza los tipos de cambio para el día siguiente.

Esta solución se puede programar para que opere en múltiples bases de datos dentro de un mismo servidor e instancia de SQL o HANA según sea el caso, tambien puede ser ejecutado en arquitecturas de 32 y 64 Bits

# Instalación
##  Requerimientos Técnicos de la Aplicación
### A nivel de Software
- Se requiere Usuario con privilegios de administrador, para configurar la tarea programada en Windows.
- Tener instalado .Net Framework 4.5.
- Credenciales de conexión a SAP BD (SQL/HANA) según sea el caso. 

La descripcion del paso a paso de la instalación de la aplicación , se indica en las paginas 7 - 11  del siguiene documento:
[Guia de Instalación](https://visualkchile.sharepoint.com/:w:/s/Desarrollo_VisualD/Efvxhu1TpmpEhNGRhflIXEcBMvyuCciX8EEioVwrN4OBOA?e=hxNLgj "Guia de Instalación")

## Documentación
[Manual de Usuario](https://visualkchile.sharepoint.com/:w:/s/Desarrollo_VisualD/Efvxhu1TpmpEhNGRhflIXEcBMvyuCciX8EEioVwrN4OBOA?e=hxNLgj "Manual de Usuario")

## Versionado
2.3.0
## Autor
Antonio Sanchez
## Licencias

