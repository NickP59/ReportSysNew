const { ipcRenderer } = require('electron');

// Установить переменную сессии
async function setSessionVariable(key, value) {
    await ipcRenderer.invoke('set-session-variable', key, value);
}

// Получить переменную сессии
async function getSessionVariable(key) {
    return await ipcRenderer.invoke('get-session-variable', key);
}

// Пример использования
setSessionVariable('EmployeeNumber', '12345');
getSessionVariable('EmployeeNumber').then(value => {
    console.log('EmployeeNumber:', value);
});
