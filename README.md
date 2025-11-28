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

## 🎞️ Galeria

![lm8UACN](https://github.com/user-attachments/assets/dc92a254-5ab9-4be8-a7c8-287fa5926869)
![M4wQuLv](https://github.com/user-attachments/assets/b4997bae-3fc7-45a6-b0d0-c35abbcb5c9c)
![MTLYnjw](https://github.com/user-attachments/assets/d724725f-5aaf-4786-a776-095fe812b41a)
<img width="1274" height="764" alt="zQWaUHS" src="https://github.com/user-attachments/assets/a5bd7853-ac4d-43d8-9260-0c563ef959d7" />
![0mmOupv](https://github.com/user-attachments/assets/dce0b903-41b9-46f3-b5fe-d8f9b3ba6c6b)
![HUrzHOy](https://github.com/user-attachments/assets/482ddef5-a4ab-4792-8e17-f839e60f9e92)
![p77KKPl](https://github.com/user-attachments/assets/adcb38b8-b6ca-4466-a030-868a6078ce55)
![O0UpgRF](https://github.com/user-attachments/assets/55af7017-6f25-40db-8f47-d1de6e17031d)
![72LFOl6](https://github.com/user-attachments/assets/6b51f52c-207d-48a9-b0a7-97e4f00c0dcd)
![HXOLlSA](https://github.com/user-attachments/assets/78b13ef6-9d27-4485-948f-d98864ace3a5)
![2PwFRhT](https://github.com/user-attachments/assets/8e6dd15b-5a41-4f01-88b6-cf2080ee301a)
![PDuNySB](https://github.com/user-attachments/assets/a4e8b770-a486-474d-94c7-30ff5d113c86)
![zvwXlCf](https://github.com/user-attachments/assets/bdb7bf19-a35f-494f-9cf0-247c6b5f8f57)
<img width="1000" height="1000" alt="RVVUhWZ" src="https://github.com/user-attachments/assets/f3551270-4ff8-498b-8295-d5f7579ef53c" />

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
