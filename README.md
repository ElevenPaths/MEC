# ExtensionControl

<img src="https://www.elevenpaths.com/wp-content/uploads/2014/03/11paths_logo.png?v=3&s=200" title="LogoApp" alt="LogoApp">

ExtensionControl (ExtProtector) es un programa que permite analizar un archivo antes de abrirlo para determinar su tipo mas allá de la extensión que posea y luego lo abre con la aplicación adecuada.  Los sistemas Windows solo permiten asociar archivos de acuerdo a su extensión, pero por medio de ExtProtector todos los archivos son abiertos con la aplicación correcta sin tener en cuenta la extensión, que puede haber sido modificada intencionalmente por malware

## Developer Mod.

En la versión actual, al ejecutarlo desde Visual Studio, la app va a modificar los registros relacionados con que Software abre por default documentos Word.
Luego, cuando intentemos abrir un documentos con extension .docx, lo primero que va a realizar es la verificacion del File Siganture (magic number) y establecer cual es el Software con el cual lo debe abrir.
Para la primera instalacion / ejecucion se necesitan permisos de administrador.


### Prerequisites

Office 365 - Office 2010

## Languajes y Technologies

* C#

## Authors

* **Jose Sperk** 
* **Sergio De los Santos** 

