# FORMULARIO-ISTG-21-v2 
<center>
  <img src="https://cdn.icon-icons.com/icons2/1156/PNG/512/1486565571-microsoft-office-excel_81549.png" alt="Excel" width="200">
</center>

Desarrollo de un pequeño formulario usando Excel con macros, VB para el control de propiedad planta y equipo.

El formulario fue diseñado para el ISTG para  mejorar el registro y control de activos 


## Software Utilizado

En este proyecto, se utilizó el software Excel con las siguientes herramientas:

- Macros
- Visual Basic
- Simulador base de datos

## Lenguaje de Programación
<center>

  <img src="https://cdn.icon-icons.com/icons2/2107/PNG/512/file_type_vb_icon_130098.png" alt="VB." width="200">
</center>

El formulario fue implementado utilizando Visual Basic que viene integrado en el Excel

## Método para insertar las imágenes en la Base

```
    Set IMG = VBA.CreateObject("Scripting.FileSystemObject")
        origen = Me.RUTA.Value
        'destino = "C:\Users\HP\Desktop\PRODUCTOS\"
        destino = "C:\Users\HP\Desktop\PRODUCTOS\" & Me.CODIGO.Value & ".jpg"
        IMG.CopyFile origen, destino
        'Range("V7").Value = origen
        Range("V7").Value = Me.CODIGO.Value & ".jpg"
       
    Else
        Range("V7").Value = "No existe imagen"
    
    End If

```
 
## Version Utilizada
Versiones superiores de 2010 de Excel, recomendado el 365



> [!CAUTION]
> Si se maneja versionas del 2010 para abajo el resultado no podría ser el correcto por la falta de algunas librerías que manejan las versiones actuales.

## Anexos
### Portada inicial
![Formulario ITSG.](https://github.com/Kevin-Saquinga/ImagenesGit/blob/main/Portada.png?raw=true)



### Pagina principal del Formulario
![Formulario ITSG.](https://github.com/Kevin-Saquinga/ImagenesGit/blob/main/Formulario.png?raw=true)



### Registrar Nuevos activos
![Formulario ITSG.](https://github.com/Kevin-Saquinga/ImagenesGit/blob/main/Registrar.png?raw=true)


### Eliminar o Modificar Activos
![Formulario ITSG.](https://github.com/Kevin-Saquinga/ImagenesGit/blob/main/EditDelet.png?raw=true)

