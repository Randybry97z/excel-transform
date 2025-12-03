const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { transformExcel } = require('./transform');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Configurar multer para almacenar archivos temporalmente
const upload = multer({
  dest: 'uploads/',
  fileFilter: (req, file, cb) => {
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/octet-stream'
    ];
    
    if (allowedMimes.includes(file.mimetype) || file.originalname.endsWith('.xlsx') || file.originalname.endsWith('.xls')) {
      cb(null, true);
    } else {
      cb(new Error('Solo se permiten archivos Excel (.xlsx, .xls)'));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB
  }
});

// Crear directorio de uploads si no existe
const ensureUploadsDir = async () => {
  try {
    await fs.mkdir('uploads', { recursive: true });
    await fs.mkdir('outputs', { recursive: true });
  } catch (error) {
    console.error('Error creando directorios:', error);
  }
};

ensureUploadsDir();

// Almacenar progreso por sesión
const progressStore = new Map();

// Endpoint para transformar Excel
app.post('/api/transform', upload.single('excelFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No se proporcionó ningún archivo' });
  }

  const inputPath = req.file.path;
  const sessionId = req.body.sessionId || Date.now().toString();
  const outputFileName = `converted_${Date.now()}.xlsx`;
  const outputPath = path.join('outputs', outputFileName);

  try {
    // Inicializar progreso
    progressStore.set(sessionId, { current: 0, total: 0, status: 'processing' });

    // Transformar con callback de progreso
    await transformExcel(inputPath, outputPath, (current, total) => {
      progressStore.set(sessionId, { current, total, status: 'processing' });
    });

    // Completar progreso
    const progress = progressStore.get(sessionId);
    if (progress) {
      progressStore.set(sessionId, { 
        ...progress, 
        status: 'completed',
        outputFile: outputFileName
      });
    }

    // Limpiar archivo temporal de entrada
    await fs.unlink(inputPath);

    res.json({
      success: true,
      message: 'Archivo transformado correctamente',
      outputFile: outputFileName,
      sessionId: sessionId
    });

  } catch (error) {
    console.error('Error al transformar:', error);
    
    // Limpiar archivos en caso de error
    try {
      await fs.unlink(inputPath);
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (outputExists) {
        await fs.unlink(outputPath);
      }
    } catch (cleanupError) {
      console.error('Error al limpiar archivos:', cleanupError);
    }

    progressStore.set(sessionId, { status: 'error', error: error.message });
    res.status(500).json({ error: error.message || 'Error al transformar el archivo' });
  }
});

// Endpoint para obtener progreso
app.get('/api/progress/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const progress = progressStore.get(sessionId);

  if (!progress) {
    return res.status(404).json({ error: 'Sesión no encontrada' });
  }

  res.json(progress);
});

// Endpoint para descargar archivo transformado
app.get('/api/download/:filename', async (req, res) => {
  const { filename } = req.params;
  const filePath = path.join('outputs', filename);

  try {
    await fs.access(filePath);
    res.download(filePath, filename, async (err) => {
      if (err) {
        console.error('Error al descargar:', err);
        res.status(500).json({ error: 'Error al descargar el archivo' });
      } else {
        // Opcional: eliminar archivo después de descargar
        // await fs.unlink(filePath);
      }
    });
  } catch (error) {
    res.status(404).json({ error: 'Archivo no encontrado' });
  }
});

// Limpiar archivos antiguos periódicamente (cada hora)
setInterval(async () => {
  try {
    const uploads = await fs.readdir('uploads');
    const outputs = await fs.readdir('outputs');
    const now = Date.now();
    const maxAge = 60 * 60 * 1000; // 1 hora

    for (const file of [...uploads, ...outputs]) {
      const filePath = path.join(file.startsWith('converted_') ? 'outputs' : 'uploads', file);
      const stats = await fs.stat(filePath);
      if (now - stats.mtimeMs > maxAge) {
        await fs.unlink(filePath);
      }
    }
  } catch (error) {
    console.error('Error en limpieza de archivos:', error);
  }
}, 60 * 60 * 1000);

app.listen(PORT, () => {
  console.log(`Servidor ejecutándose en http://localhost:${PORT}`);
});
