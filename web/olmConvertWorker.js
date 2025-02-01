importScripts("https://cdn.jsdelivr.net/pyodide/v0.27.2/full/pyodide.js");

var pyodide = null;
var olmFile = null;
var includeAttachments = true;

async function main() {
    pyodide = await loadPyodide();
    // Pyodide is now ready to use...
    await pyodide.runPythonAsync(`
        from pyodide.http import pyfetch
        response = await pyfetch("olmConvertWeb.py")
        with open("olmConvertWeb.py", "wb") as f:
            f.write(await response.bytes())
    `);
    olmConvertWeb = pyodide.pyimport("olmConvertWeb");
    await olmConvertWeb.init();

    postMessage("ready");
};

async function convert() {
    if (olmFile == undefined) {
        postMessage("fileunselected");
        return;
    }
    olmFileBytes = await olmFile.bytes();
    await pyodide.runPythonAsync(`
        from js import olmConvertWeb
        olmConvertWeb.convert()
    `);
    console.log("Finished, now going to read file...");
    let readFile = pyodide.FS.readFile("/outputEmls.zip");
    console.log("Now going to post file...");
    postMessage(readFile);
}

main();

onmessage = (e) => {
    olmFile = e.data.file;
    includeAttachments = e.data.includeAttachments;
    convert();
}