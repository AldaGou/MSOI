# **Instalador de Office LTSC - PowerShell Script**

Este script en PowerShell está diseñado para facilitar la instalación de las versiones de **Office LTSC** de forma rápida y profesional. Incluye configuraciones personalizadas para elegir la versión, el idioma, y las aplicaciones adicionales como Project y Visio.

---

## **Características**
- 🚀 **Interfaz interactiva**: Menú dinámico para seleccionar la versión, el idioma, y los productos adicionales.
- 🔧 **Exclusión de aplicaciones**: Posibilidad de excluir aplicaciones innecesarias como Bing, Teams y OneDrive.
- 🌍 **Compatibilidad multilenguaje**: Configura el idioma de la instalación (es-ES, en-US, fr-FR, etc.).
- 💾 **Automatización total**: Descarga y configuración de Office Deployment Tool (ODT) automatizada.
- ✅ **Permisos de administrador**: Asegura que se ejecuta con los privilegios necesarios para evitar problemas.

---

## **Requisitos**
1. **Sistema Operativo**: Windows 10 o superior.
2. **PowerShell**: Versión 5.1 o superior.
3. **Permisos de administrador**.
4. **Conexión a Internet** para descargar las herramientas necesarias.

---

## **Instrucciones de Uso**

1. **Descargar el script**  
   Descarga este repositorio en tu computadora:
   ```bash
   git clone https://github.com/TU-USUARIO/office-ltsc-installer.git
   cd office-ltsc-installer
   ```

2. **Ejecutar el script**  
   Abre PowerShell como administrador y ejecuta el script:
   ```bash
   powershell -File Install-OfficeLTSC.ps1
   ```

3. **Seguir las instrucciones**  
   El script te guiará a través de un menú interactivo donde podrás:
   - Seleccionar la versión de Office LTSC: 2024, 2021, o 2019.
   - Elegir el idioma (español, inglés, etc.).
   - Incluir productos adicionales como Project o Visio.

4. **Instalación**  
   El script descargará automáticamente las herramientas necesarias, generará el archivo de configuración, y comenzará la instalación.

---

## **Ejemplo del Menú Interactivo**
```
==========================================
          Office LTSC Installer           
==========================================
Configurador interactivo para Office LTSC
Por favor, siga las instrucciones.

1. Office LTSC 2024
2. Office LTSC 2021
3. Office LTSC 2019
Seleccione la versión de Office LTSC: _
```

---

## **Configuración Personalizada**
El script permite modificar el archivo de configuración generado (`configuration.xml`) para personalizar aún más la instalación. Por defecto, excluye aplicaciones como Bing y Teams.

Si deseas incluir/excluir más aplicaciones, edita la sección de exclusión:
```xml
<ExcludeApp ID="Groove" />
<ExcludeApp ID="OneDrive" />
```

---

## **Contribuciones**
¡Las contribuciones son bienvenidas! Si tienes sugerencias o mejoras, crea un issue o un pull request en este repositorio.

---

## **Licencia**
Este proyecto está bajo la licencia [MIT](LICENSE). Siéntete libre de usarlo y adaptarlo según tus necesidades.

---

## **Contacto**
Si tienes preguntas o problemas, no dudes en abrir un issue o contactarme en [tu-email@dominio.com](mailto:tu-email@dominio.com).
