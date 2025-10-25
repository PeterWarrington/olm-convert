importScripts("https://cdn.jsdelivr.net/pyodide/v0.27.2/full/pyodide.js");

var pyodide = null;
var olmFile = null;
var includeAttachments = true;
var timestamps = false;
var format = "eml";

async function main() {
    pyodide = await loadPyodide();
    // Pyodide is now ready to use...
    await pyodide.runPythonAsync(`
        from pyodide.http import pyfetch
        response = await pyfetch("olmConvertWeb.py?v=2.1")
        with open("olmConvertWeb.py", "wb") as f:
            f.write(await response.bytes())
    `);
    olmConvertWeb = pyodide.pyimport("olmConvertWeb");
    await pyodide.loadPackage("python-dateutil")
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
        pyodide.runPythonAsync(`
            from js import olmConvertWeb
            olmConvertWeb.convert()
        `).then(() => {
            let readFile = pyodide.FS.readFile("/outputEmls.zip");
            if (readFile.length > 23) {
                postMessage(readFile);
                postMessage("complete");
            } else {
                postMessage("error:Unable to convert this olm file.");
            }
        });
    };
    fileReader.readAsArrayBuffer(olmFile);
}

main();

onmessage = (e) => {
    if (typeof e.data === 'string' || e.data instanceof String) {
        if (e.data.startsWith("getfileinzip:") || e.data.startsWith("getfile:")) {
            if (e.data.startsWith("getfileinzip:")) {
                pyodide.runPython(`
                import zipfile
                archive = zipfile.ZipFile('/outputEmls.zip', 'r')
                archive.extract('${e.data.split(":")[1]}')
                `);
            }
            let decoder = new TextDecoder('utf8');
            let fileText = decoder.decode(pyodide.FS.readFile(e.data.split(":")[1]));
            postMessage("fileresponse:" + fileText);
        }
    } else {
        olmFile = e.data.file;
        includeAttachments = e.data.includeAttachments;
        format = e.data.format;
        convert();
    }
}