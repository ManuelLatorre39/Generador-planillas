//////////////////////////
// CONFIGURACIÓN CLAVE
//////////////////////////
// 1. REEMPLAZA ESTO CON EL ID DE CLIENTE WEB Q DEBERÁS CREAR EN GOOGLE CLOUD
// (Tipo: Aplicación Web, con Origenes en tu dominio github.io)
const CLIENT_ID = '29342078628-4i53r4765f7gmvvmiiguaqv6im702g9c.apps.googleusercontent.com';
const API_KEY = ''; // NO NECESITAS API_KEY PARA PEOPLE API CON USUARIOS AUTENTICADOS, SOLO OAUTH. Dejar vacío o quitar si gapi lo permite sin quejarse.

// Alcances para People Directory
const DISCOVERY_DOC = 'https://www.googleapis.com/discovery/v1/apis/people/v1/rest';
const SCOPES = 'https://www.googleapis.com/auth/directory.readonly';

// Contraseña básica de acceso al portal 
// (Cambiar aquí a la que desees; cuidado: es evaluada en cliente asique es pública para quien lee el código JS)
const APP_PASSWORD = 'LCC';

//////////////////////////
// CÓDIGO BÁSICO DE NAVEGACIÓN
//////////////////////////
let tokenClient;
let gapiInited = false;
let gisInited = false;
let finalContactsData = []; 

function logMessage(message) {
    const consoleDiv = document.getElementById('log-console');
    const p = document.createElement('div');
    p.textContent = message;
    consoleDiv.appendChild(p);
    consoleDiv.scrollTop = consoleDiv.scrollHeight;
}

function checkPassword() {
    const pwd = document.getElementById("app-password").value;
    if (pwd === APP_PASSWORD) {
        document.getElementById('section-login').classList.remove('active');
        document.getElementById('section-google-auth').classList.add('active');
        logMessage("Acesso portal concedido.");
    } else {
        document.getElementById('login-error').style.display = 'block';
    }
}

//////////////////////////
// INICIALIZACIÓN GOOGLE API
//////////////////////////
function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}
document.querySelector('script[src="https://apis.google.com/js/api.js"]').onload = gapiLoaded;

async function initializeGapiClient() {
    await gapi.client.init({
        discoveryDocs: [DISCOVERY_DOC],
    });
    gapiInited = true;
    checkIfReady();
}

function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: '', // defined at request time
    });
    gisInited = true;
    checkIfReady();
}
document.querySelector('script[src="https://accounts.google.com/gsi/client"]').onload = gisLoaded;

function checkIfReady() {
    if (gapiInited && gisInited) {
        logMessage("Google APIs iniciadas.");
    }
}

function handleAuthClick() {
    tokenClient.callback = async (resp) => {
        if (resp.error !== undefined) {
            throw (resp);
        }
        document.getElementById('section-google-auth').classList.remove('active');
        document.getElementById('section-process').classList.add('active');
        logMessage("Sesión de Google iniciada corectamente.");
    };

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({prompt: 'consent'});
    } else {
        tokenClient.requestAccessToken({prompt: ''});
    }
}

//////////////////////////
// LÓGICA CORE EXCEL Y QUERIES
//////////////////////////
function normalizeName(name, level = 0) {
    let normalizedName = name;
    
    if (level === 0) {
        if (name.includes(',')) {
            const [lastName, firstNames] = name.split(', ', 2);
            const firstNamesWords = firstNames.split(' ');
            const firstTwoNames = firstNamesWords.slice(0, 2).join(' ');
            normalizedName = `${firstTwoNames} ${lastName}`;
        }
    } else if (level === 1) {
        if (name.includes(',')) {
            const [lastName, firstNames] = name.split(', ', 2);
            const firstName = firstNames.split(' ')[0] || '';
            normalizedName = `${firstName} ${lastName}`;
        }
    }

    // Remover tildes y caracteres especiales
    normalizedName = normalizedName.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    normalizedName = normalizedName.replace(/[^\w\s]/gi, "");
    normalizedName = normalizedName.replace(/\s+/g, ' ').trim();

    return normalizedName;
}

// Función asincrónica para delay
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function processFile() {
    const fileInput = document.getElementById('file-input');
    if (fileInput.files.length === 0) {
        alert("Por favor selecciona un archivo primero.");
        return;
    }
    
    document.getElementById('btn-process').disabled = true;
    logMessage("Leyendo archivo...");

    const file = fileInput.files[0];
    const reader = new FileReader();

    // SheetJS carga CSV, XLS, XLSX directo y unificado
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[firstSheetName];
            
            // Transformar a Array de Strings (matriz)
            const rows = XLSX.utils.sheet_to_json(sheet, {header: 1});
            
            let headerIdx = -1;
            for (let i = 0; i < rows.length; i++) {
                const rowUpper = rows[i].map(c => String(c).toUpperCase().trim());
                if (rowUpper.some(c => c.includes('ALUMNO')) && rowUpper.some(c => c.includes('LEGAJO'))) {
                    headerIdx = i;
                    break;
                }
            }

            if (headerIdx === -1) {
                logMessage("Error: No se localizaron las columnas 'ALUMNO' y 'LEGAJO'.");
                document.getElementById('btn-process').disabled = false;
                return;
            }

            // Mapear headers reales
            const headers = rows[headerIdx].map(h => String(h).toUpperCase().trim());
            const indexAlumno = headers.findIndex(h => h.includes('ALUMNO'));
            const indexLegajo = headers.findIndex(h => h.includes('LEGAJO'));

            // Construir JSON mapeado con alumno y legajo
            const parsedData = [];
            for (let i = headerIdx + 1; i < rows.length; i++) {
                if (rows[i].length === 0 || !rows[i][indexAlumno]) continue;
                parsedData.push({
                    'ALUMNO': rows[i][indexAlumno],
                    'LEGAJO': rows[i][indexLegajo] || 'Sin definir'
                });
            }

            logMessage(`Alumnos detectados: ${parsedData.length}`);
            logMessage(`Iniciando cruce de datos con la API. No cierre la pestaña...`);
            
            finalContactsData = [];
            await fetchFromGoogle(parsedData);
            
            logMessage(`-------------------------------------------------`);
            logMessage(`✅ ¡Proceso finalizado con éxito!`);
            document.getElementById('btn-download').style.display = 'inline-block';
            document.getElementById('btn-process').hidden = true;

        } catch (err) {
            logMessage(`Error procesando archivo: ${err}`);
            document.getElementById('btn-process').disabled = false;
        }
    };
    reader.readAsArrayBuffer(file);
}

async function fetchFromGoogle(students) {
    for (let student of students) {
        let isFound = false;

        // Level 0
        isFound = await searchAndAppendPersona(student, 0);
        
        // Pausa preventiva cuota (Arox. 85 peticiones minuto max 90)
        await delay(750);

        // Level 1 si no se encontró
        if (!isFound) {
            isFound = await searchAndAppendPersona(student, 1);
            await delay(750);
        }

        // Failsafe format array '-'
        if (!isFound) {
            finalContactsData.push({
                'LEGAJO': student['LEGAJO'], 
                'NOMBRE': student['ALUMNO'], 
                'EMAIL': '-'
            });
        }
    }
}

async function searchAndAppendPersona(student, normLevel) {
    const queryName = normalizeName(student['ALUMNO'], normLevel);
    try {
        const response = await gapi.client.people.people.searchDirectoryPeople({
            query: queryName,
            sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
            readMask: 'names,emailAddresses'
        });

        const people = response.result.people || [];
        if (people.length > 0) {
            const firstPerson = people[0];
            const display_name = (firstPerson.names && firstPerson.names.length > 0) ? firstPerson.names[0].displayName : student['ALUMNO'];
            const email = (firstPerson.emailAddresses && firstPerson.emailAddresses.length > 0) ? firstPerson.emailAddresses[0].value : 'No Email';
            
            finalContactsData.push({
                'LEGAJO': student['LEGAJO'],
                'NOMBRE': display_name,
                'EMAIL': email
            });
            logMessage(`+ [OK] ${display_name} -> ${email}`);
            return true; // Encontrado
        }
        return false; // No encontrado
    } catch (err) {
        if (err.status === 429) {
            logMessage(`⚠️ Advertencia: Limite de cuota 429 alcanzado. Pausando de emergencia por 10 segundos...`);
            await delay(10000);
            
            // Reintenta esta iteración
            return await searchAndAppendPersona(student, normLevel);
        } else {
            logMessage(`X Error buscando a ${student['ALUMNO']}: ${err.result ? err.result.error.message : err.status}`);
            return false;
        }
    }
}

//////////////////////////
// DESCARGA FINAL 
//////////////////////////
function downloadExcel() {
    if (finalContactsData.length === 0) return;
    
    // Crear workbook base a partir de nuestro array
    const worksheet = XLSX.utils.json_to_sheet(finalContactsData);
    const workbook = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(workbook, worksheet, "Contactos");
    
    // Forzar descarga local (.xlsx)
    XLSX.writeFile(workbook, "listado_con_mails_generado.xlsx");
}