const worker = new Worker("olmConvertWorker.js");
let started = false;

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

function getFormat() {
    if (document.getElementById("html-option").checked)
        return document.getElementById("html-option").value
    else if (document.getElementById("eml-option").checked)
        return document.getElementById("eml-option").value
}

function previewDownload() {
    document.getElementById("file-preview").src;
}

function newFileHandler(filename) {
    if (getFormat() == "html") {
        $("#output-files").append(`<option value="${filename}">${filename.split("/").at(-1)}</option>`);
        $("#output-files").selectpicker('refresh');
        $("#output-files").empty();
        document.getElementsByClassName("dropdown-toggle")[0].click();
    }
}

function getFile(filename) {
    worker.postMessage(`getfile:${filename}`);
}

worker.onmessage = (e) => {
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
            let html = new TextDecoder().decode(Uint8Array.fromBase64(e.data.slice(e.data.indexOf(":") + 1)));
            let iframe = document.getElementById("file-preview");
            iframe.src = `about:blank`;
            iframe.contentWindow.document.open();
            iframe.contentWindow.document.write(html);
            iframe.contentWindow.document.close();

            let download_btn = document.getElementById("preview-download");
            download_btn.href = `data:text/html,<!DOCTYPE HTML>
                                <html>
                                <head>${iframe.contentWindow.document.head.innerHTML}</head>
                                <body>${iframe.contentWindow.document.body.innerHTML}</body>
                                <html>`;
            download_btn.download = document.getElementById("output-files").value.split("/").at(-1);
        } else if (e.data == "fileunselected") {
            document.getElementById("status-text").innerText = "You must select an OLM file.";
        } else if (e.data == "createzip") {
            document.getElementById("status-text").innerText = "Creating zip file...";
            document.getElementById("progress-bar").classList.add("bg-success");
        } else if (e.data == "startconvert") {
            document.getElementById("status-text").innerText = "Converting OLM contents to EML files...";
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
            document.getElementById("progress-bar").classList.remove("bg-success");
            document.getElementById("status-text").innerText = "Complete!";
            started = false;
        }
    } else {
        let blob = new Blob([e.data], {
            type: "application/zip"
        });

        document.getElementById("preview-btn").classList.remove("d-none");
        (new bootstrap.Offcanvas('#preview-offcanvas')).show();
    
        let url = window.URL.createObjectURL(blob);

        let download_btn = document.getElementById("download-btn");
        download_btn.href = url;
        download_btn.download = "emlOutput.zip";
        download_btn.classList.remove("d-none");
    }
}