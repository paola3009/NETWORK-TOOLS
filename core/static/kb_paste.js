/**
 * kb_paste.js
 * Barra de Ferramentas de Formatação (Negrito, Itálico, Sublinhado, Cores, Cabeçalhos, Tabelas)
 * + Colar imagens diretamente da área de transferência (Print Screen / Ctrl+V)
 * + Arrastar e soltar (drag & drop) arquivos de imagem no campo de texto Markdown da Knowledge Base.
 */
document.addEventListener('DOMContentLoaded', function () {
    const textarea = document.getElementById('conteudo_md');
    if (!textarea) return;

    // Helper para inserir texto na posição do cursor
    function insertAtCursor(textToInsert) {
        const start = textarea.selectionStart;
        const end = textarea.selectionEnd;
        const val = textarea.value;
        textarea.value = val.substring(0, start) + textToInsert + val.substring(end);
        textarea.selectionStart = textarea.selectionEnd = start + textToInsert.length;
        textarea.focus();
    }

    // Helper para envolver texto selecionado com prefixo e sufixo
    function wrapSelection(prefix, suffix, defaultText = "texto") {
        const start = textarea.selectionStart;
        const end = textarea.selectionEnd;
        const val = textarea.value;
        const selected = val.substring(start, end);

        if (selected) {
            textarea.value = val.substring(0, start) + prefix + selected + suffix + val.substring(end);
            textarea.selectionStart = start + prefix.length;
            textarea.selectionEnd = end + prefix.length;
        } else {
            textarea.value = val.substring(0, start) + prefix + defaultText + suffix + val.substring(end);
            textarea.selectionStart = start + prefix.length;
            textarea.selectionEnd = start + prefix.length + defaultText.length;
        }
        textarea.focus();
    }

    // --- BARRA DE FERRAMENTAS DE FORMATAÇÃO ---
    document.querySelectorAll('.kb-tool-btn').forEach(btn => {
        btn.addEventListener('click', function (e) {
            e.preventDefault();
            const format = btn.getAttribute('data-format');
            const color = btn.getAttribute('data-color');

            if (color) {
                wrapSelection(`<font color="${color}">`, `</font>`, "Texto colorido");
                return;
            }

            switch (format) {
                case 'bold':
                    wrapSelection('**', '**', 'texto em negrito');
                    break;
                case 'italic':
                    wrapSelection('*', '*', 'texto em itálico');
                    break;
                case 'underline':
                    wrapSelection('<u>', '</u>', 'texto sublinhado');
                    break;
                case 'mark':
                    wrapSelection('<mark>', '</mark>', 'texto destacado');
                    break;
                case 'h1':
                    wrapSelection('\n# ', '\n', 'Título Principal');
                    break;
                case 'h2':
                    wrapSelection('\n## ', '\n', 'Seção / Subtítulo');
                    break;
                case 'h3':
                    wrapSelection('\n### ', '\n', 'Subseção');
                    break;
                case 'bullet':
                    wrapSelection('\n- ', '\n', 'Item com marcador');
                    break;
                case 'number':
                    wrapSelection('\n1. ', '\n', 'Item numerado');
                    break;
                case 'code':
                    wrapSelection('\n```bash\n', '\n```\n', '# Comandos de rede ou script aqui');
                    break;
                case 'table':
                    insertAtCursor('\n| Coluna 1 | Coluna 2 | Coluna 3 |\n| --- | --- | --- |\n| Item 1 | Item 2 | Item 3 |\n');
                    break;
                case 'callout':
                    wrapSelection('\n> [!NOTE]\n> ', '\n', 'Aviso importante ou recomendação técnica');
                    break;
            }
        });
    });

    // Seletor de Cores via Dropdown
    const colorSelector = document.getElementById('kb-font-color-select');
    if (colorSelector) {
        colorSelector.addEventListener('change', function () {
            const chosenColor = colorSelector.value;
            if (chosenColor) {
                wrapSelection(`<font color="${chosenColor}">`, `</font>`, "Texto colorido");
                colorSelector.value = ""; // Reseta o select
            }
        });
    }

    // --- UPLOAD E INSERÇÃO DE IMAGENS ---
    function uploadAndInsertImage(blob, originalName) {
        const timestamp = new Date().toISOString().replace(/[-:T.]/g, '').slice(0, 14);
        const filename = originalName || `print_${timestamp}.png`;
        const placeholder = `\n![Enviando imagem: ${filename}...]()\n`;
        
        insertAtCursor(placeholder);

        const formData = new FormData();
        formData.append('image', blob, filename);

        fetch('/kb/upload_image', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) throw new Error('Falha HTTP ' + response.status);
            return response.json();
        })
        .then(data => {
            if (data.success && data.url) {
                const markdownTag = `\n![${data.nome || filename}](${data.url})\n`;
                textarea.value = textarea.value.replace(placeholder, markdownTag);
            } else {
                throw new Error('Resposta sem sucesso');
            }
        })
        .catch(err => {
            console.warn('Upload AJAX falhou, usando fallback em Base64:', err);
            const reader = new FileReader();
            reader.onload = function (evt) {
                const base64Tag = `\n![${filename}](${evt.target.result})\n`;
                textarea.value = textarea.value.replace(placeholder, base64Tag);
            };
            reader.readAsDataURL(blob);
        });
    }

    // 1. Captura evento de Colar (Ctrl + V / Área de Transferência)
    textarea.addEventListener('paste', function (e) {
        const clipboard = e.clipboardData || window.clipboardData;
        if (!clipboard || !clipboard.items) return;

        const items = clipboard.items;
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            if (item.type && item.type.startsWith('image/')) {
                e.preventDefault();
                const blob = item.getAsFile();
                if (blob) {
                    uploadAndInsertImage(blob, 'screenshot.png');
                }
                return;
            }
        }
    });

    // 2. Drag & Drop de imagens diretamente no campo de texto
    textarea.addEventListener('dragover', function (e) {
        e.preventDefault();
        textarea.classList.add('drag-active');
    });

    textarea.addEventListener('dragleave', function (e) {
        textarea.classList.remove('drag-active');
    });

    textarea.addEventListener('drop', function (e) {
        textarea.classList.remove('drag-active');
        const files = e.dataTransfer ? e.dataTransfer.files : null;
        if (!files || files.length === 0) return;

        let hasImage = false;
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (file.type && file.type.startsWith('image/')) {
                if (!hasImage) {
                    e.preventDefault();
                    hasImage = true;
                }
                uploadAndInsertImage(file, file.name);
            }
        }
    });

    // 3. Botão auxiliar de Upload de Imagem
    const btnUploadImg = document.getElementById('btn_inserir_imagem');
    const inputImgPicker = document.getElementById('input_picker_imagem');
    if (btnUploadImg && inputImgPicker) {
        btnUploadImg.addEventListener('click', function () {
            inputImgPicker.click();
        });
        inputImgPicker.addEventListener('change', function () {
            if (inputImgPicker.files && inputImgPicker.files[0]) {
                const file = inputImgPicker.files[0];
                uploadAndInsertImage(file, file.name);
                inputImgPicker.value = '';
            }
        });
    }
});
