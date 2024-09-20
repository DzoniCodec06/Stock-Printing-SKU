const { app, BrowserWindow } = require("electron");

const createWin = () => {
    let win = new BrowserWindow({
        width: 700,
        height: 400,
        maximizable: false,
        resizable: false,
        autoHideMenuBar: true,
        webPreferences: {
            devTools: true,
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    win.loadFile("./src/index.html");
    //win.webContents.openDevTools();
}

app.on("ready", () => {
    createWin();

    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) createWin();
    })
});

app.on("window-all-closed", () => {
    if (process.platform != "darwin") app.quit();
})