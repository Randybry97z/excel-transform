# Conversor de Excel

Aplicación web para transformar archivos Excel con una interfaz gráfica moderna.

## Características

- 📤 Carga de archivos Excel mediante arrastrar y soltar o selección
- 📊 Barra de progreso en tiempo real durante la transformación
- ⬇️ Descarga del archivo transformado
- 🎨 Interfaz oscura moderna y responsive

## Instalación

1. Instala las dependencias:
```bash
npm install
```

## Uso

1. Inicia el servidor:
```bash
npm start
```

2. Abre tu navegador en: `http://localhost:3000`

3. Selecciona o arrastra un archivo Excel (.xlsx o .xls)

4. Haz clic en "Transformar Excel" y espera a que se complete el proceso

5. Descarga el archivo transformado cuando esté listo

## Estructura del Proyecto

```
excel-conversion/
├── transform.js      # Lógica de transformación de Excel
├── server.js         # Servidor Express
├── package.json      # Dependencias del proyecto
├── public/           # Archivos estáticos (interfaz web)
│   ├── index.html    # Página principal
│   ├── style.css     # Estilos
│   └── app.js        # Lógica del frontend
├── uploads/           # Archivos temporales de entrada (se crea automáticamente)
└── outputs/          # Archivos transformados (se crea automáticamente)
```

## Dependencias

- **express**: Servidor web
- **multer**: Manejo de archivos subidos
- **exceljs**: Procesamiento de archivos Excel
- **cors**: Habilitar CORS

## Notas

- Los archivos temporales se eliminan automáticamente después de 1 hora
- El tamaño máximo de archivo es 50MB
- El script `transform.js` también puede ejecutarse desde línea de comandos:
  ```bash
  node transform.js input.xlsx output.xlsx
  ```

