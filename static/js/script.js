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

/**NUEVO CODIGO 16/03/25
**/
document.getElementById('selectAllBtn').addEventListener('click', function() {
    document.querySelectorAll('.file-checkbox').forEach(checkbox => {
      checkbox.checked = document.getElementById('generateCheckbox').checked;
    });
  });

  document.getElementById('GenerarDocu').addEventListener('click', function() {
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
            if (file.type === "text/csv") {
                let listItem = document.createElement('li');
                listItem.textContent = `${file.name} - ${Math.round(file.size / 1024)} KB`;
                uploadList.appendChild(listItem);
            } else {
                alert('Solo se permiten archivos CSV');
            }
        });
    }
});

// Obtener todos los elementos con la clase opGen
const opGenElements = document.querySelectorAll('.opGen');

// Agregar un evento click a cada elemento
opGenElements.forEach(element => {
    element.addEventListener('click', () => {
        // Remover la clase active de todos los elementos
        opGenElements.forEach(el => el.classList.remove('active'));

        // Agregar la clase active al elemento clickeado
        element.classList.add('active');
    });
});
