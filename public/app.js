const fileInput = document.getElementById('fileInput');
const uploadLabel = document.getElementById('uploadLabel');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const removeBtn = document.getElementById('removeBtn');
const transformBtn = document.getElementById('transformBtn');
const downloadBtn = document.getElementById('downloadBtn');
const progressSection = document.getElementById('progressSection');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const progressDetails = document.getElementById('progressDetails');
const message = document.getElementById('message');

let selectedFile = null;
let sessionId = null;
let progressInterval = null;

// Drag and drop
uploadLabel.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadLabel.classList.add('dragover');
});

uploadLabel.addEventListener('dragleave', () => {
    uploadLabel.classList.remove('dragover');
});

uploadLabel.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadLabel.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelect(files[0]);
    }
});

// File input change
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

// Remove file
removeBtn.addEventListener('click', () => {
    selectedFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    transformBtn.disabled = true;
    hideMessage();
});

// Transform button
transformBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    // Reset UI
    transformBtn.disabled = true;
    downloadBtn.style.display = 'none';
    hideMessage();
    progressSection.style.display = 'block';
    progressBar.style.width = '0%';
    progressText.textContent = '0%';
    progressDetails.textContent = 'Iniciando transformación...';

    // Generate session ID
    sessionId = Date.now().toString();

    // Create form data
    const formData = new FormData();
    formData.append('excelFile', selectedFile);
    formData.append('sessionId', sessionId);

    try {
        // Start progress polling
        startProgressPolling();

        // Send file to server
        const response = await fetch('/api/transform', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.error || 'Error al transformar el archivo');
        }

        // Stop progress polling
        stopProgressPolling();

        // Show success
        progressBar.style.width = '100%';
        progressText.textContent = '100%';
        progressDetails.textContent = 'Transformación completada';
        
        showMessage('Archivo transformado correctamente', 'success');
        
        // Show download button
        downloadBtn.style.display = 'block';
        downloadBtn.onclick = () => {
            window.location.href = `/api/download/${result.outputFile}`;
        };

    } catch (error) {
        stopProgressPolling();
        progressSection.style.display = 'none';
        showMessage(error.message || 'Error al transformar el archivo', 'error');
        transformBtn.disabled = false;
    }
});

function handleFileSelect(file) {
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'application/octet-stream'
    ];
    
    const isValidType = validTypes.includes(file.type) || 
                       file.name.endsWith('.xlsx') || 
                       file.name.endsWith('.xls');

    if (!isValidType) {
        showMessage('Por favor, selecciona un archivo Excel (.xlsx o .xls)', 'error');
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
    transformBtn.disabled = false;
    downloadBtn.style.display = 'none';
    progressSection.style.display = 'none';
    hideMessage();
}

function startProgressPolling() {
    progressInterval = setInterval(async () => {
        if (!sessionId) return;

        try {
            const response = await fetch(`/api/progress/${sessionId}`);
            const progress = await response.json();

            if (progress.status === 'processing') {
                const percentage = progress.total > 0 
                    ? Math.round((progress.current / progress.total) * 100) 
                    : 0;
                progressBar.style.width = `${percentage}%`;
                progressText.textContent = `${percentage}%`;
                progressDetails.textContent = `Procesando fila ${progress.current} de ${progress.total}...`;
            } else if (progress.status === 'completed') {
                stopProgressPolling();
                progressBar.style.width = '100%';
                progressText.textContent = '100%';
                progressDetails.textContent = 'Transformación completada';
            } else if (progress.status === 'error') {
                stopProgressPolling();
                throw new Error(progress.error || 'Error en el procesamiento');
            }
        } catch (error) {
            console.error('Error al obtener progreso:', error);
        }
    }, 500); // Poll every 500ms
}

function stopProgressPolling() {
    if (progressInterval) {
        clearInterval(progressInterval);
        progressInterval = null;
    }
}

function showMessage(text, type) {
    message.textContent = text;
    message.className = `message ${type}`;
    message.style.display = 'block';
}

function hideMessage() {
    message.style.display = 'none';
}

