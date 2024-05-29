document.getElementById('fileInput').addEventListener('change', handleFileSelect);

function handleFileSelect(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const arrayBuffer = e.target.result;
        const fileType = file.type;

        if (fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            convertXlsxToHtml(arrayBuffer);
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            convertDocxToHtml(arrayBuffer);
        } else if (fileType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
            convertPptxtoHtml(arrayBuffer);
        } else {
            alert('Formato de arquivo não suportado!');
        }
    };

    reader.readAsArrayBuffer(file);
}

function convertDocxToHtml(arrayBuffer) {
    mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
        .then(result => {
            const htmlContent = result.value;
            document.getElementById('filePreviewContainer').innerHTML = htmlContent;
        })
        .catch(error => {
            alert('Erro ao converter DOCX para HTML: ' + error.message);
            console.error('Erro ao converter DOCX para HTML:', error);
        });
}

function convertXlsxToHtml(arrayBuffer) {
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    let tableHtml = '<table><tr>';

    // Cabeçalho da tabela
    sheetData[0].forEach(header => {
        tableHtml += '<th>' + header + '</th>';
    });
    tableHtml += '</tr>';

    // Linhas da tabela
    sheetData.slice(1).forEach(row => {
        tableHtml += '<tr>';
        row.forEach(cell => {
            tableHtml += '<td>' + (cell !== undefined ? cell : '') + '</td>';
        });
        tableHtml += '</tr>';
    });
    tableHtml += '</table>';

    document.getElementById('filePreviewContainer').innerHTML = tableHtml;
}

function convertPptxtoHtml(arrayBuffer) {
    const pptx = new PptxGenJS();
    pptx.load(arrayBuffer, 'arraybuffer').then(() => {
        let htmlContent = '';

        pptx.slides.forEach(slide => {
            htmlContent += '<div class="slide">';
            slide.slideObjects.forEach(obj => {
                if (obj.text) {
                    htmlContent += `<p>${obj.text}</p>`;
                } else if (obj.image) {
                    htmlContent += `<img src="${obj.image.src}" alt="Slide image"/>`;
                }
            });
            htmlContent += '</div>';
        });

        document.getElementById('filePreviewContainer').innerHTML = htmlContent;
    });
}

function convertHtmlToPdf(htmlContent) {
    const element = document.createElement('div');
    element.innerHTML = htmlContent;

    html2pdf()
        .from(element)
        .toPdf()
        .output('arraybuffer')
        .then(pdfArrayBuffer => {
            const blob = new Blob([pdfArrayBuffer], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);

            // Mostrar o PDF na tela
            const object = document.createElement('object');
            object.data = url;
            object.type = 'application/pdf';
            object.width = '100%';
            object.height = '600px';
            document.getElementById('filePreviewContainer').innerHTML = '';
            document.getElementById('filePreviewContainer').appendChild(object);
        })
        .catch(error => {
            alert('Erro ao converter para PDF: ' + error.message);
            console.error('Erro ao converter para PDF:', error);
        });
}

function convertFileToPdf() {
    const filePreviewContainer = document.getElementById('filePreviewContainer');
    const htmlContent = filePreviewContainer.innerHTML;
    convertHtmlToPdf(htmlContent);
}
