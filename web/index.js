const worker = new Worker("olmConvertWorker.js");
const textEncoder = new TextEncoder();
let started = false;

let filenameList = [];

document.getElementById("convertBtn").disabled = true;
document.getElementById("convertBtn").onclick = (e) => {
    outputList = [];
    document.getElementById("output-files").innerHTML = "";

    let iframe = document.getElementById("file-preview");
    iframe.src = `about:blank`;
    iframe.contentWindow.document.open();
    iframe.contentWindow.document.write(`<!DOCTYPE HTML><b style="margin: 10px; font-family: sans-serif">Select a file to preview it here...</b>`);
    iframe.contentWindow.document.close();

    let cnvBtn = document.getElementById("convertBtn");
        cnvBtn.disabled = true;
        cnvBtn.innerText = "Converting...";
    worker.postMessage({
        file: document.getElementById("olmFile").files[0],
        includeAttachments: document.getElementById("includeAttachments").checked,
        format: getFormat()
    });
};

document.getElementById('changelog-btn').addEventListener('click', () =>
    bootstrap.Toast.getOrCreateInstance(
        document.getElementById('changelog-toast')
    ).show()
);

document.getElementById("preview-print").addEventListener("click", () => {
    document.getElementById("file-preview").contentWindow.print();
});

function getFormat() {
    if (document.getElementById("html-option").checked)
        return document.getElementById("html-option").value
    else if (document.getElementById("eml-option").checked)
        return document.getElementById("eml-option").value
}

function newFileHandler(filename) {
    if (getFormat() == "html") {
        filenameList.push(filename);
    }
}

function getFileInZip(filename) {
    worker.postMessage(`getfileinzip:${filename}`);
}

function fileFilter(value) {
    document.getElementById("output-files").innerHTML = "";
    let filterList = filenameList.filter(str => value.split(" ").every(
        keyword => str.toLowerCase().includes(keyword.toLowerCase())
    )).slice(0,200);

    filterList.forEach(filename => {
        let option = document.createElement("option");
        option.text = filename.split("/").at(-1);
        option.value = filename;
        document.getElementById("output-files").add(option);
    });

    if (filterList.length >= 200) {
        let option = document.createElement("option");
        option.text = "Filter to see more...";
        option.value = "NONE";
        document.getElementById("output-files").add(option);
    }

}

worker.onmessage = async (e) => {
    if (typeof e.data === 'string' || e.data instanceof String) {
        if (e.data.startsWith("progress:")) {
            if (!started) {
                if (getFormat() == "html") {} else {
                    document.getElementById("preview-btn").classList.add("d-none");
                }
                started = true;
            }
            let percent = e.data.slice(e.data.indexOf(":") + 1);
            let progressbar = document.getElementById("progress-bar");
            progressbar.style.width = percent;
            progressbar.innerText = percent;
        } else if (e.data.startsWith("newfile:")) {
            let filename = e.data.slice(e.data.indexOf(":") + 1);
            newFileHandler(filename);
        } else if (e.data.startsWith("fileresponse:")) {
            let html = e.data.slice(e.data.indexOf(":") + 1);
            let iframe = document.getElementById("file-preview");
            iframe.src = `about:blank`;
            iframe.contentWindow.document.open();
            iframe.contentWindow.document.write(html);
            iframe.contentWindow.document.close();

            let download_btn = document.getElementById("preview-download");
            download_btn.href = "data:text/html;base64," + await bufferToBase64(textEncoder.encode(`<!DOCTYPE HTML>
                                <html>
                                <head>
                                ${iframe.contentWindow.document.head.innerHTML}
                                <style>body {font-family: sans-serif}</style>
                                </head>
                                <body>${iframe.contentWindow.document.body.innerHTML}</body>
                                <html>`));
            download_btn.download = document.getElementById("output-files").value.split("/").at(-1);
        } else if (e.data == "fileunselected") {
            document.getElementById("status-text").innerText = "You must select an OLM file.";
        } else if (e.data == "createzip") {
            document.getElementById("status-text").innerText = "Creating zip file...";
            document.getElementById("progress-bar").classList.add("bg-success");
        } else if (e.data == "startconvert") {
            document.getElementById("status-text").innerText = "Converting OLM contents...";
        } else if (e.data == "fileread") {
            document.getElementById("status-text").innerText = "Reading OLM file...";
        }

        if (e.data == "ready" || e.data == "complete" || e.data.startsWith("error:") || e.data == "fileunselected") {
            let cnvBtn = document.getElementById("convertBtn");
            cnvBtn.disabled = false;
            cnvBtn.innerText = "Convert!";
        }

        if (e.data.startsWith("error:")) {
            document.getElementById("status-text").innerText = "An error occurred: " + e.data.slice(e.data.indexOf(":") + 1);
        }

        if (e.data == "complete") {
            if (getFormat() == "html") {
                fileFilter("");

                document.getElementById("preview-btn").classList.remove("d-none");
                (new bootstrap.Offcanvas('#preview-offcanvas')).show();

                document.getElementById("output-files").addEventListener("change", (event) => {
                    getFileInZip(event.target.value);
                });

                let outputSearchChange = () => fileFilter(document.getElementById("output-search").value);
                document.getElementById("output-search").onchange = outputSearchChange;
                document.getElementById("output-search").onkeydown = outputSearchChange;
            }

            document.getElementById("progress-bar").classList.remove("bg-success");
            document.getElementById("status-text").innerText = "Complete!";
            started = false;
        }
    } else {
        let blob = new Blob([e.data], {
            type: "application/zip"
        });

        let url = window.URL.createObjectURL(blob);

        let download_btn = document.getElementById("download-btn");
        download_btn.href = url;
        download_btn.download = "emlOutput.zip";
        download_btn.classList.remove("d-none");
    }
}

// https://stackoverflow.com/a/66046176/5270231
async function bufferToBase64(buffer) {
    // use a FileReader to generate a base64 data URI:
    const base64url = await new Promise(r => {
        const reader = new FileReader()
        reader.onload = () => r(reader.result)
        reader.readAsDataURL(new Blob([buffer]))
    });
    // remove the `data:...;base64,` part from the start
    return base64url.slice(base64url.indexOf(',') + 1);
}