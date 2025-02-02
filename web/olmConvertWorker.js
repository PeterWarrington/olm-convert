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
    let fileReader = new FileReader();
    fileReader.onload = async function() {
        olmFileBytes = this.result;
        await pyodide.runPythonAsync(`
            from js import olmConvertWeb
            olmConvertWeb.convert()
        `);
        let readFile = pyodide.FS.readFile("/outputEmls.zip");
        if (readFile.length > 23) {
            postMessage(readFile);
            postMessage("complete");
        } else {
            postMessage("error:Unable to convert this olm file.");
        }
    };
    fileReader.readAsArrayBuffer(olmFile);
}

main();

onmessage = (e) => {
    olmFile = e.data.file;
    includeAttachments = e.data.includeAttachments;
    convert();
}