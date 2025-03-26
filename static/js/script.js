// Globals
const formOptions = ['asignacionesTXT', 'PAEG_TySI', 'asignaciones1TXT', 'asignaciones2TXT'];
let siteConfigs = {
    form: 1, // 1: asignacionesTXT | 2: PAEG_TySI | 3: asignaciones1TXT | 4: asignaciones2TXT
}

// Función para mostrar/ocultar secciones
/** 
document.addEventListener('DOMContentLoaded', function () {
    // Función para mostrar un div específico y ocultar los demás
    window.mostrarDiv1 = function (divId) {
        // Oculta todos los divs con la clase 'div-oculto'
        //// const divs = document.querySelectorAll('.div-oculto');
        //// divs.forEach(div => {
        ////     div.style.display = 'none';
        //// });
        const acercaSection = document.getElementById('ace

        // Muestra el div específico
        const divToShow = document.getElementById(divId);
        if (divToShow) {
            divToShow.style.display = 'block';
        }
    };

    // Otra función similar si se requiere mostrar diferentes secciones
    window.mostrarDiv = function (divId) {
        // Oculta todos los divs que podrían mostrarse
        const allDivs = document.querySelectorAll('.page');
        allDivs.forEach(div => {
            div.style.display = 'none';
        });

        // Muestra el div seleccionado
        const targetDiv = document.getElementById(divId);
        if (targetDiv) {
            targetDiv.style.display = 'block';
        }
    };

    // Event listeners para el menú de navegación
    document.querySelectorAll('.menu-item').forEach(item => {
        item.addEventListener('click', function (event) {
            event.preventDefault();
            const page = this.getAttribute('data-page');
            if (page === 'page1') {
                mostrarDiv1('acerca');
            } else if (page === 'page2') {
                mostrarDiv1('documentos');
            } else if (page === 'page3') {
                mostrarDiv1('plantillas_section');
            }
        });
    });
});
*/
function addHandlerToSelectAllBTN() {
    document.getElementById('selectAllBtn')?.addEventListener('click', function () {
        document.querySelectorAll('.file-checkbox').forEach(checkbox => {
            checkbox.checked = document.getElementById('generateCheckbox').checked;
        });
    });

    document.getElementById('GenerarDocu')?.addEventListener('click', function () {
        let selectedFiles = [];
        document.querySelectorAll('.file-checkbox:checked').forEach(checkbox => {
            selectedFiles.push(checkbox.value);
        });

        if (selectedFiles.length === 0) {
            alert('Por favor, seleccione al menos un archivo.');
            return;
        }

        fetch("{{ url_for('generar_GruposPT') }}", {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ archivos: selectedFiles })
        })
            .then(response => response.json())
            .then(data => alert('Documentos generados correctamente'))
            .catch(error => console.error('Error:', error));
    });
}





//FUNCION PARA APERECER DESAPARECER LOS ELEMENTOS PRINCIPALES DE LA VENTANA DE VISUALIZACION AL USUARIO
function mostrarDiv1(divId) {
    // Oculta todos los divs con la clase 'page'
    document.querySelectorAll('.page').forEach(div => {
        div.classList.add('hidden');
    });

    // Muestra el div seleccionado
    const targetDiv = document.getElementById(divId);
    if (targetDiv) {
        targetDiv.classList.remove('hidden');
    }
}



//FUNCION PARA CARGAR ARCHIVOS Y SE EJECUTE EL PROGRAMA
document.addEventListener('DOMContentLoaded', function () {
    const dropZone = document.getElementById('drop');
    const fileInput = document.getElementById('file-input-1');
    const uploadList = document.getElementById('upload-list');
    getConfigs();
    checkFormIdFromQuery();

    // Prevenir comportamientos por defecto
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // Añadir estilos al arrastrar
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('highlight'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('highlight'));
    });

    // Manejo del evento drop
    dropZone.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const files = e.dataTransfer.files;
        handleFiles(files);
    }

    // Abrir el diálogo de archivos al hacer clic
    dropZone.addEventListener('click', () => fileInput.click());

    // Manejo del input file
    fileInput.addEventListener('change', (e) => {
        const files = e.target.files;
        handleFiles(files);
    });

    function handleFiles(files) {
        [...files].forEach(file => {
            console.log('file', file);
            console.log('file.type', file.type);
            if (["text/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "text/plain", "application/vnd.ms-excel"].includes(file.type)) {
                let listItem = document.createElement('li');
                listItem.textContent = `${file.name} - ${Math.round(file.size / 1024)} KB`;
                uploadList.appendChild(listItem);
            } else {
                alert('Solo se permiten archivos CSV, XLSX, XLS, TXT');
            }
        });
    }
    handleOpGenClick();
    addHandlerToSelectAllBTN();
});

// Configs
function getConfigs() {
    const configsInLocal = localStorage.getItem('config_UTILTECH_4685');
    console.log('configsInLocal', configsInLocal);
    if (configsInLocal !== null) {
        const configsInLocalParsed = JSON.parse(configsInLocal);
        console.log('configsInLocalParsed', configsInLocalParsed);
        console.log('typeof configsInLocalParsed', typeof configsInLocalParsed);
        if (configsInLocalParsed !== null && typeof configsInLocalParsed === 'object') {
            Object.keys(siteConfigs).forEach(key => {
                if (configsInLocalParsed[key] !== undefined) {
                    siteConfigs[key] = configsInLocalParsed[key];
                }
            });
        }
    }
    console.log('siteConfigs', siteConfigs);
    configFormOption(siteConfigs.form);
}

// Change form option
function changeFormOption(formOption) {
    siteConfigs.form = formOption;
    localStorage.setItem('config_UTILTECH_4685', JSON.stringify(siteConfigs));
    chargeUrlWithNewQuery();
}

// configFormOption
function configFormOption(formOption) {
    const formInput = document.getElementById('file-input-1');
    formInput.name = formOptions[formOption - 1];
    formInput.accept = formOption === 1 ? '.csv' : formOption === 2 ? '.xlsx,.xls' : '.txt';
    const formOptionInput = document.getElementById('file-input-option-1');
    formOptionInput.value = 'formulario' + formOption;
    const formOption2Checkbox = document.getElementById('file-checkbox-option-2');
    formOption2Checkbox.value = 'formulario' + formOption;
    const formMessageInput = document.getElementById('file-input-message-1');
    formMessageInput.innerHTML = formOption === 1 ? 'Sube el archivo ConfirmacionDeAsesoriasPT en formato CSV' : formOption === 2 ? 'Sube el archivo propuesta en formato XLSX o XLS' : 'Por favor, seleccione un archivo TXT';
}

// Crear Function handler para los opGen
function handleOpGenClick() {
    // Obtener todos los elementos con la clase opGen
    const opGenElements = document.querySelectorAll('.opGen');

    // Agregar un evento click a cada elemento
    opGenElements.forEach(element => {
        element.addEventListener('click', () => {
            console.log('element', element);
            // Remover la clase active de todos los elementos
            opGenElements.forEach(el => el.classList.remove('active'));

            // Agregar la clase active al elemento clickeado
            element.classList.add('active');
            let formOption = parseInt(element.innerHTML.split('.')[0]);
            changeFormOption(formOption);
        });
    });
}
/* 
    Crear una función que haga lo siguiente
    1. Revisar si hay query en el url.
    2. Si hay query, revisar si existe la query form_id.
    3. Si existe la query form_id, revisar si el valor es "formulario" más la concatenación de un número entre 1 y 4.
    4. Si el valor es válido, cambiar el form en siteConfigs si es difertente al que ya se tiene.
    5. Si no existe la query form_id, revisar el form en siteConfigs y si tiene un valor diferente a 1, cargar de nuevo el sitio pero añadiendo el query para tener el form_id correspondiente a "formulario" + siteConfigs.form.
*/
function checkFormIdFromQuery() {
    const urlParams = new URLSearchParams(window.location.search);
    const formId = urlParams.get('form_id');
    if (formId !== null) {
        const formIdNumber = parseInt(formId.split('formulario')[1]);
        if (formIdNumber >= 1 && formIdNumber <= 4) {
            if (siteConfigs.form !== formIdNumber) {
                changeFormOption(formIdNumber);
            }
        }
    } else {
        chargeUrlWithNewQuery();
    }
}

function chargeUrlWithNewQuery(){
    console.log('siteConfigs.form', siteConfigs);
    if (window.location.pathname === '/') {
        const url = new URL(window.location.href);
        url.searchParams.set('form_id', `formulario${siteConfigs.form}`);
        window.location.href = url.toString();
    }
}
