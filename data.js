const courseData = [
    {
        day: 1,
        title: "Introducción a la Ofimática y Microsoft Word — Fundamentos",
        tool: "Word",
        objective: "Reconocer el entorno de herramientas ofimáticas, configurar el espacio de trabajo en Word y crear documentos con formato básico profesional.",
        lessons: [
            {
                title: "¿Qué es la ofimática? Aplicaciones en la Empresa",
                content: "La ofimática es el conjunto de técnicas, aplicaciones y herramientas informáticas que se utilizan en funciones de oficina para optimizar, automatizar y mejorar los procedimientos o tareas relacionados. <br><br><b>En el entorno administrativo:</b> Se usa para la creación de informes, gestión de bases de datos de clientes, control de inventarios y comunicación interna. <br><b>En el entorno empresarial:</b> Permite la toma de decisiones basada en datos, la colaboración remota y la estandarización de procesos legales y comerciales.",
                videoUrl: "https://www.tiktok.com/embed/v2/7535721632464145669",
                tip: "La ofimática moderna ya no es solo local; hoy se vive en la nube (Cloud Computing) para trabajar desde cualquier lugar."
            },
            {
                title: "Panorama del paquete Microsoft 365 / Office",
                content: "Microsoft 365 es una suite integrada que ofrece herramientas específicas para cada necesidad empresarial: <br>• <b>Microsoft Word:</b> Procesador de textos para cartas, contratos y manuales. <br>• <b>Microsoft Excel:</b> Hojas de cálculo para contabilidad, presupuestos y análisis de datos. <br>• <b>Microsoft PowerPoint:</b> Diseño de presentaciones visuales para ventas o reportes. <br>• <b>Outlook/Teams:</b> Comunicación y gestión de tiempo.",
                videoUrl: "https://www.youtube.com/embed/tRXegIr_BJc?start=63",
                tip: "Regla de oro: Si tiene mucho texto, usa Word. Si tiene muchos números, usa Excel. Si es para una audiencia, usa PowerPoint."
            },
            {
                title: "Entorno de Microsoft Word: La Interfaz Profesional",
                content: "Dominar la interfaz es clave para la velocidad: <br>1. <b>Cinta de opciones:</b> Organizada por pestañas (Inicio, Insertar, Disposición) que agrupan todas las funciones. <br>2. <b>Barra de acceso rápido:</b> Ubicada en la parte superior izquierda; puedes personalizarla para tener 'Guardar' o 'Deshacer' siempre a un clic. <br>3. <b>Vistas y Paneles:</b> Modo lectura, Diseño de impresión y Diseño web. El Panel de Navegación (Ctrl + B) es vital para moverse en documentos largos.",
                imageUrl: "assets/entorno-word.png",
                tip: "Haz clic derecho en cualquier herramienta de la cinta y elige 'Agregar a la barra de acceso rápido' para ahorrar segundos valiosos."
            },
            {
                title: "Gestión de Archivos: Guardado Local vs OneDrive",
                content: "El flujo de trabajo profesional implica saber dónde vive tu archivo: <br>• <b>Crear y Abrir:</b> Uso de plantillas vs documentos en blanco. <br>• <b>Guardar (Local):</b> Almacenamiento en el disco duro del PC (.docx). <br>• <b>Guardar (OneDrive):</b> Almacenamiento en la nube que permite el Autoguardado y la coedición. <br>• <b>Exportar a PDF:</b> Crucial para documentos que no deben ser modificados (facturas, contratos finales).",
                tip: "Si activas el 'Autoguardado' en la esquina superior izquierda (solo en OneDrive), nunca más perderás trabajo por un apagón."
            },
            {
                title: "Formato de Texto: Estética y Legibilidad",
                content: "Un documento administrativo debe ser legible y serio: <br>• <b>Fuente y Tamaño:</b> Tipografías 'Sans Serif' como Calibri o Arial (11-12 pts) son el estándar. <br>• <b>Énfasis:</b> Usa <b>Negrita</b> para títulos, <i>Cursiva</i> para términos extranjeros y <u>Subrayado</u> solo si es estrictamente necesario. <br>• <b>Color y Resaltado:</b> Úsalos con moderación para marcar puntos críticos o pendientes de revisión.",
                videoUrl: "https://www.youtube.com/embed/8XmN95w60M0",
                tip: "Evita usar más de dos tipos de fuentes diferentes en un mismo documento administrativo para mantener el profesionalismo."
            },
            {
                title: "Formato de Párrafo: Estructura del Documento",
                content: "La organización visual del texto: <br>• <b>Alineación:</b> Justificado (Ctrl+J) para informes formales; Izquierda para cartas modernas. <br>• <b>Interlineado:</b> El espacio entre líneas. 1.15 es el estándar; 2.0 es para borradores o normas APA. <br>• <b>Sangría y Espaciado:</b> La sangría de primera línea ayuda a identificar párrafos. El espaciado entre párrafos evita que el texto se vea como un bloque sólido.",
                videoUrls: [
                    "https://www.youtube.com/embed/Tqu7l18Xjio",
                    "https://www.youtube.com/embed/dl8z49u84Yg"
                ],
                tip: "Usa el botón 'Mostrar todo' (¶) en la pestaña Inicio para ver espacios, saltos de página y caracteres ocultos que descuadran tu documento."
            }
        ],
        practice: "<b>Laboratorio 1:</b> Crea un documento nuevo, configura márgenes de 2.54cm, usa fuente Arial 11. Redacta un memorando con alineación justificada e interlineado 1.15. Guárdalo como PDF.",
        quiz: {
            question: "¿Qué herramienta usarías para encontrar rápidamente una palabra en un documento de 50 páginas?",
            options: ["Panel de Navegación (Ctrl+B)", "Barra de Acceso Rápido", "Cinta de Opciones", "Formato de Párrafo"],
            correct: 0
        },
        mediaPlaceholder: { type: "video", url: "https://www.youtube.com/embed/59h7_G9sI9o", description: "Interfaz de Word 365." }
    },
    {
        day: 2,
        title: "Microsoft Word — Documentos Profesionales",
        tool: "Word",
        objective: "Dominar herramientas de edición avanzada para producir documentos con estructura y navegación profesional.",
        lessons: [
            {
                title: "Uso de Estilos",
                content: "Los estilos permiten generar Tablas de Contenido automáticas y mantener la jerarquía del documento (Título 1, Título 2).",
                tip: "Modifica un estilo desde el panel derecho y se aplicará a todo el documento."
            },
            {
                title: "Tablas y Columnas",
                content: "Inserción y diseño de tablas con bordes profesionales y uso de columnas para boletines informativos.",
                tip: "Usa 'Ajustar texto' para que las imágenes no rompan el formato de tus párrafos."
            },
            {
                title: "Encabezados y Pies de Página",
                content: "Configuración de logos en la parte superior y numeración automática en la inferior.",
                tip: "Usa 'Saltos de Sección' si quieres que la numeración empiece después de la portada."
            }
        ],
        practice: "Elaborar un informe administrativo de 2 páginas con portada y tabla de contenido.",
        quiz: {
            question: "¿Qué permite generar una tabla de contenidos automática?",
            options: ["Los Estilos", "Las Negritas", "El Guardado", "Los Márgenes"],
            correct: 0
        },
        mediaPlaceholder: { type: "image", url: "https://www.ablebits.com/office-addins-blog/wp-content/uploads/2020/02/word-table-of-contents-example.png", description: "Estructura de estilos en Word." }
    },
    {
        day: 3,
        title: "Microsoft Excel — Fundamentos y Fórmulas Esenciales",
        tool: "Excel",
        objective: "Comprender la estructura de celdas y aplicar fórmulas básicas para cálculos administrativos.",
        lessons: [
            {
                title: "Entorno y Tipos de Datos",
                content: "Diferencia entre Texto, Número, Fecha y Moneda. Importancia del formato de celda.",
                tip: "Asegúrate de que las fechas tengan formato de 'Fecha corta' para poder operar con ellas."
            },
            {
                title: "Funciones Vitales: SUMA, PROMEDIO y CONTAR",
                content: "Uso de funciones para automatizar sumatorias de gastos, promedios de ventas y conteo de registros.",
                tip: "Toda función en Excel debe empezar obligatoriamente con el signo igual (=)."
            },
            {
                title: "Referencias Absolutas ($)",
                content: "Uso del signo $ para fijar celdas en fórmulas cuando arrastramos hacia abajo o hacia los lados.",
                tip: "Pulsa F4 mientras escribes la fórmula para alternar los signos de dólar automáticamente."
            }
        ],
        practice: "Construir una planilla de control de gastos mensuales.",
        quiz: {
            question: "¿Qué signo se usa para fijar una celda en Excel?",
            options: ["$", "#", "&", "%"],
            correct: 0
        },
        mediaPlaceholder: { type: "video", url: "https://www.youtube.com/embed/A_G26K0Xf0g", description: "Referencias en Excel." }
    },
    {
        day: 4,
        title: "Microsoft Excel — Tablas, Gráficos y Análisis",
        tool: "Excel",
        objective: "Organizar datos en tablas estructuradas y crear visualizaciones gráficas efectivas.",
        lessons: [
            {
                title: "Tablas Oficiales (Ctrl + T)",
                content: "Convertir un rango de datos en Tabla permite usar filtros automáticos y filas de totales sin fórmulas manuales.",
                tip: "Las Tablas se expanden solas cuando agregas una nueva fila al final."
            },
            {
                title: "Función SI (Lógica)",
                content: "Toma de decisiones lógicas: SI(condicion; valor_si_verdadero; valor_si_falso). Ej: Si ventas > 1000, 'Meta cumplida'.",
                tip: "Usa el Asistente de Funciones (fx) para no perderte con los paréntesis."
            },
            {
                title: "Gráficos Estadísticos",
                content: "Creación de gráficos de barras para comparar categorías y circulares para porcentajes.",
                tip: "Quita las líneas de cuadrícula para que tu gráfico se vea mucho más limpio."
            }
        ],
        practice: "Crear una tabla dinámica simple y un gráfico de barras comparativo.",
        quiz: {
            question: "¿Atajo para crear una Tabla en Excel?",
            options: ["Ctrl + T", "Ctrl + G", "Ctrl + B", "Ctrl + N"],
            correct: 0
        },
        mediaPlaceholder: { type: "image", url: "https://trumpexcel.com/wp-content/uploads/2017/03/Slicers-in-Excel.png", description: "Segmentadores en Excel." }
    },
    {
        day: 5,
        title: "Microsoft PowerPoint — Presentaciones Efectivas",
        tool: "PowerPoint",
        objective: "Diseñar presentaciones claras y coherentes con la identidad institucional.",
        lessons: [
            {
                title: "Diseño y Jerarquía Visual",
                content: "Uso de temas y variantes. Aplicación de la regla 6x6 (6 líneas por 6 palabras) para no saturar al público.",
                tip: "Menos es más. Usa imágenes de alta calidad en lugar de mucho texto."
            },
            {
                title: "Insertar Gráficos Vinculados",
                content: "Cómo traer un gráfico de Excel a PowerPoint de modo que, si el Excel cambia, el gráfico de la presentación también.",
                tip: "Usa 'Pegado especial > Pegar vínculo' para mantener la conexión viva."
            },
            {
                title: "Transiciones vs Animaciones",
                content: "Transiciones para el cambio entre diapositivas y Animaciones para los elementos internos. Uso moderado y profesional.",
                tip: "Usa 'Transformación' (Morph) para movimientos suaves y profesionales."
            }
        ],
        practice: "Crear una presentación de 8 diapositivas sobre resultados trimestrales.",
        quiz: {
            question: "¿En qué consiste la regla 6x6?",
            options: ["6 palabras por 6 líneas", "6 colores por 6 fotos", "6 segundos por 6 diapositivas", "6 temas por 6 minutos"],
            correct: 0
        },
        mediaPlaceholder: { type: "video", url: "https://www.youtube.com/embed/J_jZ5yZ3_lY", description: "Vincular Excel y PowerPoint." }
    },
    {
        day: 6,
        title: "Gestión de Correo Profesional — Outlook | GMail",
        tool: "Outlook",
        objective: "Administrar eficientemente la comunicación corporativa y normas de etiqueta digital.",
        lessons: [
            {
                title: "Firma y Redacción Profesional",
                content: "Estructura de un correo formal: Asunto claro, Saludo, Cuerpo breve, Despedida y Firma institucional.",
                tip: "Nunca envíes un correo sin asunto; puede terminar en SPAM."
            },
            {
                title: "Reglas de Bandeja de Entrada",
                content: "Creación de carpetas automáticas para que los correos de tu jefe o de ciertos clientes se clasifiquen solos.",
                tip: "Usa el 'Archivo' para mantener tu bandeja principal en cero."
            },
            {
                title: "Calendarios y Reuniones",
                content: "Programación de citas y envío de invitaciones a reuniones mediante Teams o Zoom integrado.",
                tip: "Usa el 'Asistente de programación' para ver cuándo tus compañeros están libres."
            }
        ],
        practice: "Configurar una firma corporativa y crear reglas de clasificación.",
        quiz: {
            question: "¿Qué significa CCO?",
            options: ["Copia Oculta", "Copia Oficial", "Correo Obligatorio", "Copia Ordinaria"],
            correct: 0
        },
        mediaPlaceholder: { type: "image", url: "https://blog.hubspot.es/hubfs/ejemplos-firmas-correo-electr%C3%B3nico-6.jpg", description: "Firma profesional." }
    },
    {
        day: 7,
        title: "OneDrive y Trabajo Colaborativo",
        tool: "Nube",
        objective: "Utilizar herramientas de almacenamiento para compartir y coeditar documentos de forma segura.",
        lessons: [
            {
                title: "Almacenamiento en OneDrive",
                content: "Sincronización de archivos entre tu computadora y la nube. Acceso desde cualquier dispositivo.",
                tip: "Si pierdes tu PC, tus archivos están a salvo en tu cuenta de Microsoft."
            },
            {
                title: "Compartir con Permisos",
                content: "Diferencia entre 'Solo ver' y 'Puede editar'. Uso de contraseñas y fechas de expiración en links.",
                tip: "Deja de enviar archivos por correo; envía el enlace de OneDrive."
            },
            {
                title: "Coedición en Tiempo Real",
                content: "Cómo ver los cursores de otros compañeros trabajando en el mismo documento de Word o Excel simultáneamente.",
                tip: "Usa los 'Comentarios' y las 'Menciones (@)' para comunicarte dentro del archivo."
            }
        ],
        practice: "Subir un informe a OneDrive y compartirlo con permiso de edición.",
        quiz: {
            question: "¿Qué permite la coedición?",
            options: ["Editar varias personas a la vez", "Borrar archivos", "Solo leer", "Cambiar la contraseña"],
            correct: 0
        },
        mediaPlaceholder: { type: "video", url: "https://www.youtube.com/embed/L_86Cq7S8zE", description: "Colaboración en tiempo real." }
    },
    {
        day: 8,
        title: "Excel Avanzado — Tablas Dinámicas",
        tool: "Excel",
        objective: "Analizar grandes volúmenes de datos con tablas dinámicas y automatización.",
        lessons: [
            {
                title: "Creación de Tablas Dinámicas",
                content: "Resumir miles de filas en segundos por categorías (Ventas por mes, Gastos por departamento).",
                tip: "Mantén tus datos de origen limpios (sin filas en blanco) para evitar errores."
            },
            {
                title: "Gráficos Dinámicos",
                content: "Gráficos que cambian automáticamente cuando filtras o mueves los campos de la tabla dinámica.",
                tip: "Usa 'Segmentadores' para crear botones de filtro interactivos."
            },
            {
                title: "Validación de Datos",
                content: "Crear listas desplegables para que el usuario solo pueda elegir opciones predefinidas.",
                tip: "Evita errores de dedo obligando al usuario a seleccionar de una lista."
            }
        ],
        practice: "Construir una tabla dinámica de nómina y calcular antigüedades.",
        quiz: {
            question: "¿Para qué sirve una Tabla Dinámica?",
            options: ["Resumir grandes datos", "Hacer dibujos", "Escribir cartas", "Cambiar el color"],
            correct: 0
        },
        mediaPlaceholder: { type: "image", url: "https://excelchamps.com/wp-content/uploads/2020/06/dashboard-example.png", description: "Dashboard en Excel." }
    },
    {
        day: 9,
        title: "Formularios y Documentos Legales",
        tool: "Word/Excel",
        objective: "Crear plantillas reutilizables y automatizar envíos masivos.",
        lessons: [
            {
                title: "Controles de Contenido (Formularios)",
                content: "Creación de cuadros de texto, selectores de fecha y listas desplegables dentro de Word.",
                tip: "Usa la pestaña 'Programador' (o Desarrollador) para activar estas herramientas."
            },
            {
                title: "Combinar Correspondencia",
                content: "Generar 100 cartas personalizadas en segundos uniendo una lista de Excel con un documento de Word.",
                tip: "Ideal para generar diplomas, contratos masivos o estados de cuenta."
            },
            {
                title: "Proteger Documentos",
                content: "Cómo poner contraseña a un archivo y restringir la edición para que solo se puedan llenar los formularios.",
                tip: "Usa firmas digitales básicas para dar validez a tus documentos internos."
            }
        ],
        practice: "Crear una plantilla de contrato con campos de control de contenido.",
        quiz: {
            question: "¿Qué pestaña se necesita para crear formularios?",
            options: ["Programador", "Inicio", "Vista", "Ayuda"],
            correct: 0
        },
        mediaPlaceholder: { type: "video", url: "https://www.youtube.com/embed/6c1p2KJ7H2E", description: "Combinar correspondencia." }
    },
    {
        day: 10,
        title: "Proyecto Integrador Final",
        tool: "Proyecto",
        objective: "Integrar todas las competencias en un proyecto administrativo completo.",
        lessons: [
            {
                title: "El Informe Maestro",
                content: "Creación del documento final: Informe en Word con gráficos de Excel vinculados y presentación PPT.",
                tip: "Asegúrate de que todos tus archivos estén en la misma carpeta de OneDrive."
            },
            {
                title: "Revisión y Feedback",
                content: "Repaso de dudas finales y entrega de los archivos finales mediante enlace compartido.",
                tip: "Revisa siempre la ortografía y el formato final antes de entregar."
            }
        ],
        practice: "PROYECTO: Generar el informe de gestión completo integrando todas las herramientas.",
        quiz: {
            question: "¿Cuál es el peso del proyecto final?",
            options: ["50%", "10%", "30%", "100%"],
            correct: 0
        },
        mediaPlaceholder: { type: "image", url: "https://marketplace.canva.com/EAFIE7fX8A0/1/0/1600w/canva-azul-y-dorado-moderno-certificado-de-reconocimiento-UToO5-T8p04.jpg", description: "Certificado final." }
    }
];

const evaluationCriteria = [
    { component: "Participación y asistencia", weight: "10%", description: "Asistencia a las 10 sesiones y participación activa en clase." },
    { component: "Actividades prácticas diarias", weight: "30%", description: "Entrega de los ejercicios desarrollados en cada sesión (días 1–9)." },
    { component: "Proyecto integrador final", weight: "50%", description: "Entregable completo en Word, Excel, PowerPoint y OneDrive (día 10)." },
    { component: "Autoevaluación", weight: "10%", description: "Reflexión escrita sobre los aprendizajes y áreas de mejora." }
];

export { courseData, evaluationCriteria };
