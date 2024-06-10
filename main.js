const { ipcRenderer } = require('electron');

// ���������� ���������� ������
async function setSessionVariable(key, value) {
    await ipcRenderer.invoke('set-session-variable', key, value);
}

// �������� ���������� ������
async function getSessionVariable(key) {
    return await ipcRenderer.invoke('get-session-variable', key);
}

// ������ �������������
setSessionVariable('EmployeeNumber', '12345');
getSessionVariable('EmployeeNumber').then(value => {
    console.log('EmployeeNumber:', value);
});
