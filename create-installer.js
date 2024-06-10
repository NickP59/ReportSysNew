const electronInstaller = require('electron-winstaller');
const path = require('path');

async function createInstaller() {
  try {
    await electronInstaller.createWindowsInstaller({
      appDirectory: path.join(__dirname, 'path/to/your/app'),
      outputDirectory: path.join(__dirname, 'path/to/output/directory'),
      authors: 'My App Inc.',
      exe: 'myapp.exe'
    });
    console.log('It worked!');
  } catch (e) {
    console.error(`No dice: ${e.message}`);
  }
}

createInstaller();
