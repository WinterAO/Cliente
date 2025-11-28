# WinterAO Resurrection - Cliente

![Visual Basic](https://img.shields.io/badge/Visual%20Basic-6.0-blue)
![DirectX](https://img.shields.io/badge/DirectX-8-green)
![License](https://img.shields.io/badge/license-custom-orange)

Cliente oficial del proyecto **WinterAO Resurrection**, un MMORPG 2D desarrollado en Visual Basic 6.0 con DirectX 8.

## 📋 Descripción

Este repositorio contiene el código fuente del cliente del juego WinterAO Resurrection, una mod del clásico juego Argentum Online. El cliente está desarrollado en Visual Basic 6.0 y utiliza DirectX 8 para el renderizado gráfico y la gestión de sonido.

## ✨ Características Principales

- **Motor gráfico DirectX 8**: Renderizado 2D optimizado con soporte para efectos visuales
- **Sistema de red asíncrono**: Comunicación TCP cliente-servidor eficiente
- **Sistema de partículas**: Efectos visuales dinámicos y personalizables
- **Iluminación**: Sistema de luces con soporte para ambiente nocturno/diurno
- **Sistema de clima**: Lluvia, nieve y efectos atmosféricos
- **Interfaz gráfica completa**: 
  - Gestión de inventario
  - Sistema de comercio
  - Sistema de clanes (guilds)
  - Sistema de party
  - Sistema de quests
  - Estadísticas de personaje
  - Minimapa
  - Mundo continuo
- **Sistema de habilidades**: Interface para skills y trabajos (herrero, carpintero, etc.)
- **Chat multicanal**: Soporte para diferentes tipos de mensajes
- **Sistema de personajes**: Creación y personalización de personajes

## 🛠️ Requisitos Técnicos

### Para compilar el proyecto:
- **Visual Basic 6.0** (IDE completo)
- **DirectX 8 SDK**
- **Windows XP o superior** (recomendado Windows 7/10 con modo compatibilidad)

### Dependencias incluidas:
- `DX8VB.DLL` - DirectX 8 para Visual Basic
- `MSCOMCTL.OCX` - Controles comunes de Microsoft
- `AAMD532.DLL` - Componente adicional
- `zlib.dll` - Compresión de datos

## 📦 Estructura del Proyecto

```
Cliente/
├── CODIGO/              # Código fuente principal
│   ├── *.frm           # Formularios de la interfaz
│   ├── *.bas           # Módulos de código
│   ├── *.cls           # Clases del proyecto
│   └── uControls/      # Controles personalizados
├── Init/               # Archivos de inicialización
├── Recursos/           # Recursos gráficos y de audio
├── Client.vbp          # Proyecto de Visual Basic
└── WinterAOResurrection.exe # Ejecutable compilado
```

## 🔧 Componentes Principales

### Módulos Core
- `General.bas` - Funciones generales del cliente
- `Protocol.bas` - Protocolo de comunicación con el servidor
- `Protocol_Write.bas` - Envío de paquetes al servidor
- `ProtocolCmdParse.bas` - Parseo de comandos del servidor
- `TileEngine.bas` - Motor de renderizado de tiles

### Motor DirectX 8
- `mDx8_Engine.bas` - Inicialización y gestión del motor DirectX
- `mDx8_Particulas.bas` - Sistema de partículas
- `mDx8_Luces.bas` - Sistema de iluminación
- `mDx8_Clima.bas` - Sistema de clima
- `mDx8_Text.bas` - Renderizado de texto

### Networking
- `clsSocket.cls` - Clase principal para conexiones TCP
- `modSocket.bas` - Gestión de sockets
- `TCP.bas` - Funciones de red

### Sistemas de Juego
- `clsGrapchicalInventory.cls` - Inventario gráfico
- `clsCustomKeys.cls` - Configuración de teclas personalizadas
- `clsSoundEngine.cls` - Motor de audio
- `mPooChar.bas` - Pool de personajes en pantalla
- `mPooMap.bas` - Pool de mapas

## 🚀 Ejecución

Para ejecutar el cliente necesitas:
1. El ejecutable compilado o el proyecto abierto en VB6
2. Los archivos DLL en el mismo directorio del ejecutable
3. La carpeta `Init/` con los archivos de configuración
4. La carpeta `Recursos/` con los gráficos y sonidos del juego
5. Conexión al servidor de WinterAO Resurrection

## 🔗 Enlaces

- [Repositorio del Servidor](https://github.com/WinterAO/Server)
- [Herramientas y recursos](https://github.com/orgs/WinterAO/repositories)

## ⚙️ Configuración

El cliente utiliza archivos de configuración en la carpeta `Init/` para:
- Configuración de gráficos y resolución
- Teclas personalizadas
- Configuración de audio
- Configuración de red (IP del servidor, puerto)

## 🐛 Problemas Conocidos

- Compatibilidad limitada con Windows 10/11 (requiere modo compatibilidad)
- DirectX 8 puede requerir instalación de runtime legacy en sistemas modernos
- Algunas funciones pueden requerir permisos de administrador

---

**Nota**: Este es un proyecto esta basado en Argentum Online. Todo el crédito original corresponde a los creadores de Argentum Online.
