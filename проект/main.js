const {app, BrowserWindow} = require('electron')


let win;
function createWindow() {
    win = new BrowserWindow({
        width: 1600,
        height: 1000,
        webPreferences: {
            nodeIntegration: true,
            enableRemoteModule: true
        }
    })
    win.on("closed", () => {
        win = null;
    })
    win.loadFile('./index.html')
    win.webContents.openDevTools();
}


app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit()
    }
})

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow()
    }
})
