document.addEventListener('DOMContentLoaded', () => {
    const navButtons = document.querySelectorAll('.nav-btn');
    const dynamicContents = document.querySelectorAll('.dynamic-content');

    const elements = {
        kronos: {
            dropZone: document.getElementById('kronos-drop-zone'),
            fileInput: document.getElementById('kronos-file-input'),
            dataDisplay: document.getElementById('kronos-data-display'),
            actionButton: document.getElementById('kronos-action-bar')
        },
        sales: {
            dropZone: document.getElementById('sales-drop-zone'),
            fileInput: document.getElementById('sales-file-input'),
            dataDisplay: document.getElementById('sales-data-display'),
            actionButton: document.getElementById('sales-action-bar')
        }
    };

    // Configura o botão "Lançar"
    document.querySelectorAll('.lancar-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const parent = e.target.closest('.action-bar');
            const type = parent.id.includes('kronos') ? 'kronos' : 'sales';

            const url = (type === 'kronos') ? '/lancar-kronos' : '/lancar-sales';
            alert(`Iniciando a automação de Reposição ${type.toUpperCase()}...`);

            fetch(url, {
                method: 'POST'
            })
            .then(response => response.json())
            .then(data => {
                alert(`Sucesso! ${data.message}`);
                console.log(data.message);
            })
            .catch(error => {
                console.error('Erro:', error);
                alert('Ocorreu um erro ao conectar com o servidor ou durante a automação.');
            });
        });
    });

    // Função para alternar o conteúdo
    navButtons.forEach(button => {
        button.addEventListener('click', () => {
            navButtons.forEach(btn => btn.classList.remove('active'));
            dynamicContents.forEach(content => content.classList.remove('active'));
            document.querySelectorAll('.action-bar').forEach(bar => bar.classList.remove('visible'));

            button.classList.add('active');
            const targetId = button.dataset.target;
            document.getElementById(targetId).classList.add('active');

            // Limpa o conteúdo quando a aba é trocada
            Object.values(elements).forEach(el => {
                const tableHead = el.dataDisplay.querySelector('thead tr');
                const tableBody = el.dataDisplay.querySelector('tbody');
                if (tableHead) tableHead.innerHTML = '';
                if (tableBody) tableBody.innerHTML = '';
                el.dataDisplay.classList.add('hidden');
                el.dataDisplay.querySelector('.loading-message').classList.add('hidden');
            });
        });
    });

    const currencyFormatter = new Intl.NumberFormat('pt-BR', {
        style: 'currency',
        currency: 'BRL'
    });

    // Função para processar o arquivo
    function processFile(file, elementConfig) {
        const { dataDisplay, actionButton } = elementConfig;

        // Busca o loadingMessage e os elementos da tabela aqui, quando já sabemos que o dataDisplay existe
        const loadingMessage = dataDisplay.querySelector('.loading-message');
        const tableHead = dataDisplay.querySelector('thead tr');
        const tableBody = dataDisplay.querySelector('tbody');

        const expectedFileName = (dataDisplay.id === 'kronos-data-display') ? 'Base.xlsx' : 'Sales.xlsx';

        if (file.name.toLowerCase() !== expectedFileName.toLowerCase()) {
            alert(`Erro: Por favor, selecione apenas o arquivo ${expectedFileName}.`);
            dataDisplay.classList.add('hidden');
            actionButton.classList.remove('visible');
            return;
        }

        dataDisplay.classList.remove('hidden');
        if (loadingMessage) loadingMessage.classList.remove('hidden');

        if (tableHead) tableHead.innerHTML = '';
        if (tableBody) tableBody.innerHTML = '';

        const formData = new FormData();
        formData.append('file', file);
        const uploadUrl = (dataDisplay.id === 'kronos-data-display') ? '/upload-kronos' : '/upload-sales';

        console.log(`Tentando enviar arquivo para a URL: ${uploadUrl}`);
        console.log(`Nome do arquivo: ${file.name}`);
        console.log(`Tipo do arquivo: ${file.type}`);

        fetch(uploadUrl, {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (loadingMessage) loadingMessage.classList.add('hidden');

            if (data.message) {
                console.log(data.message);
                if (data.data && data.headers) {
                    console.log("Dados recebidos:", data.data);
                    console.log("Cabeçalhos recebidos:", data.headers);

                    if (tableHead) tableHead.innerHTML = '';
                    if (tableBody) tableBody.innerHTML = '';

                    data.headers.forEach(header => {
                        const th = document.createElement('th');
                        th.textContent = header;
                        if (tableHead) tableHead.appendChild(th);
                    });

                    data.data.forEach(row => {
                        const newRow = document.createElement('tr');
                        data.headers.forEach(header => {
                            const td = document.createElement('td');
                            const cellValue = row[header] || '';

                            const isCurrencyHeader = header.toLowerCase().includes('valor');

                            if (isCurrencyHeader) {
                                const numericValue = parseFloat(String(cellValue).replace(',', '.'));
                                td.textContent = !isNaN(numericValue) ?
                                    currencyFormatter.format(numericValue) :
                                    String(cellValue || '');
                            } else {
                                td.textContent = String(cellValue);
                            }
                            newRow.appendChild(td);
                        });
                        if (tableBody) tableBody.appendChild(newRow);
                    });

                    actionButton.classList.remove('hidden');
                    actionButton.classList.add('visible');
                } else {
                    console.error('Dados ou cabeçalhos não encontrados na resposta do servidor.');
                    alert('Erro: Os dados do arquivo não puderam ser lidos corretamente.');
                }
            } else {
                alert(`Erro no servidor: ${data.message}`);
            }
        })
        .catch(error => {
            if (loadingMessage) loadingMessage.classList.add('hidden');
            console.error('Erro:', error);
            alert('Ocorreu um erro ao enviar o arquivo.');
        });
    }

    // Gerenciador de eventos unificado
    Object.keys(elements).forEach(type => {
        const el = elements[type];
        
        el.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            el.dropZone.classList.add('active-drag');
        });
        el.dropZone.addEventListener('dragleave', () => {
            el.dropZone.classList.remove('active-drag');
        });
        el.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            el.dropZone.classList.remove('active-drag');
            const file = e.dataTransfer.files[0];
            processFile(file, el);
        });
        el.fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            processFile(file, el);
        });
    });
});