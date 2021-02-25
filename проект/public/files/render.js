const {remote} = require('electron');
const dialog = remote.dialog;
const WIN = remote.getCurrentWindow();
const os = require('os');
const path = require('path');
let desktopDir = path.join(os.homedir(), "Desktop") + "\\";
let pointsFileBtn = document.getElementsByClassName("content-file-choose")[0];
let pointsPathBlock = document.getElementsByClassName("content-file-path")[0]
let outputDirectoryBtn = document.getElementsByClassName("content-file-choose")[1];
let outputDirectoryBlock = document.getElementsByClassName("content-file-path")[1]
outputDirectoryBlock.textContent = desktopDir
async function createChoosePathWindow(parentWindow, btn) {
    let options = {
        title: "Выбор файла, содержащего точки",
        defaultPath: desktopDir,
        buttonLabel: "Выбрать файл с точками",
        filters: [
            {name: 'Excel', extensions: ['xls', 'xlsx', 'xlsm']}
        ],
        properties: ["openFile"]
    }
    let chosePath = await dialog.showOpenDialog(parentWindow, options)
    if (!chosePath.canceled) {
        btn.textContent = chosePath.filePaths;
    }
}

async function createChangeFolderWindow(parentWindow, btn) {
    let options = {
        title: "Выбор директории для создания протоколов",
        defaultPath: desktopDir,
        buttonLabel: "Выбрать папку",
        properties: ["openDirectory"]
    }
    let chosePath = await dialog.showOpenDialog(parentWindow, options)
    if (!chosePath.canceled) {
        btn.textContent = chosePath.filePaths + "\\";
    }
}

outputDirectoryBtn.onclick = () => {
    createChangeFolderWindow(WIN, outputDirectoryBlock)
}

pointsFileBtn.onclick = () => {
    createChoosePathWindow(WIN, pointsPathBlock)
}