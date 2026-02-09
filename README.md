# Planificador Académico ESPOL

Una aplicación de escritorio para planificar horarios académicos de la Escuela Superior Politécnica del Litoral (ESPOL), con funcionalidad de extracción automática de datos del portal académico.

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/4c732a21-883a-43e5-aced-4c8191887bf9" />

## Características Principales

- **Extracción Automática de Datos**: Recopila información directamente del portal académico ESPOL con manejo integrado de CAPTCHA
- **Visualización de Horarios**: Interfaz intuitiva con cuadrícula horaria semanal (Lunes a Viernes, 7:00-21:00)
- **Detección Automática de Conflictos**: Verifica choques de horarios de clases y exámenes en tiempo real
- **Sugerencias Inteligentes**: Recomienda materias compatibles basadas en el horario actual del usuario
- **Gestión de Exámenes**: Panel dedicado para visualizar todas las fechas de exámenes programados
- **Exportación de Horarios**: Capacidad para guardar horarios como imágenes PNG

### Navegador Requerido
- Mozilla Firefox (necesario para Selenium WebDriver)

## Guía de Uso

### Extracción de Datos
- **Extracción Automática**: Seleccionar "Scraping ESPOL" e ingresar credenciales del portal académico
- **Carga Manual**: Utilizar "Cargar CSV" para importar datos previamente descargados

### Construcción de Horarios
- Seleccionar celdas en la cuadrícula para ver opciones disponibles
- Cada materia se identifica mediante etiquetas con iniciales únicas en la esquina superior derecha
- Las celdas con color de fondo representan materias ya inscritas

### Gestión Académica
- **Inscripción**: Seleccionar "INSCRIBIR" para agregar una materia al horario
- **Cambio de Paralelo**: Utilizar "CAMBIAR" para modificar el paralelo de una materia existente
- **Retiro de Materias**: Eliminar materias con la opción "RETIRAR"
- **Gestión de Conflictos**: Las materias con choques de horario muestran indicador de "CONFLICTO"

### Funcionalidades Avanzadas
- **Sistema de Filtros**: Búsqueda específica por materia, docente o día
- **Motor de Sugerencias**: Encuentra materias compatibles en el panel lateral
- **Panel de Exámenes**: Revisa fechas y aulas de exámenes programados
- **Exportación Visual**: Guarda horarios completos como imágenes con la función "Capturar"

### Manejo de CAPTCHA
La aplicación implementa un sistema híbrido para CAPTCHA:
1. Inicia automáticamente el navegador para resolución manual por el usuario
2. Espera confirmación de resolución completa
3. Continúa con la extracción de datos de manera automatizada

### Almacenamiento de Datos
- Los datos extraídos se almacenan en formato CSV (`INFORME.csv`)
- Estructura columnar optimizada para materias, horarios, docentes y exámenes

## Consideraciones Importantes

- **Seguridad de Credenciales**: Las credenciales no se almacenan permanentemente, solo se utilizan durante la sesión activa
- **Visibilidad del Navegador**: Durante la extracción, el navegador permanece visible para interacción del usuario

