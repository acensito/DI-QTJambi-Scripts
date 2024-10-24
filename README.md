# DI-QTJambi-Scripts
Scripts para poder compilar archivos .jui en QT Jambi bajo Windows 10.

## Utilizacion
Para poderlos utilizar, sigue los siguientes pasos:

- Clona o descarga el zip de este repositorio.
- Copia o mueve el directorio /salida y los archivos juic.vbs y mod.vbs a la carpeta /bin de tu instalacion de QTJambi
- Crea tus formularios .ui con QT Designer y guardalos en la carpeta /salida
- Una vez creado los formularios que te interesen ejecuta mod.vbs para transformarlos a .jui
- Creados los archivos .jui, ejecuta juic.vbs para compilarlos todos a archivos .java
- Estos archivos los podrás ya trapasar a tu proyecto java. No olvides a estos archivos indicarles en la primera linea el paquete al que pertenecen.

## Nota
- No me hago en absoluto responsable del mal uso de estos scripts. Tampoco aseguro que funcionen al 100% pero a mi me estan resultando.
- Estos scripts me han funcionado bajo Windows 10, QT Jambi 4.8.6 y 4.5.2 con Java 17 LTS y Netbeans 23.

## Mejoras
- El script juic.vbs procesará por lotes sin tanto mensaje.

- Si tienes alguna mejora, haz un pull request, serán bienvenidas!