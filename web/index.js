const worker = new Worker("olmConvertWorker.js");

document.getElementById("convertBtn").disabled = true;
document.getElementById("convertBtn").onclick = (e) => {
    let cnvBtn = document.getElementById("convertBtn");
        cnvBtn.disabled = true;
        cnvBtn.innerText = "Converting...";
    worker.postMessage({
        file: document.getElementById("olmFile").files[0],
        includeAttachments: document.getElementById("includeAttachments").checked
    });
};

worker.onmessage = (e) => {
    if (typeof e.data === 'string' || e.data instanceof String) {
        if (e.data.startsWith("progress:")) {
            let percent = e.data.slice(e.data.indexOf(":") + 1);
            let progressbar = document.getElementById("progress-bar");
            progressbar.style.width = percent;
            progressbar.innerText = percent;
        } else if (e.data == "fileunselected") {
            document.getElementById("error-text").innerText = "You must select an OLM file.";
        } else if (e.data == "ready" || e.data == "complete") {
            let cnvBtn = document.getElementById("convertBtn");
            cnvBtn.disabled = false;
            cnvBtn.innerText = "Convert!";
        } else if (e.data == "createzip") {
            document.getElementById("status-text").innerText = "Creating zip file...";
            document.getElementById("progress-bar").classList.add("bg-success");
        } else if (e.data == "startconvert") {
            document.getElementById("status-text").innerText = "Converting OLM contents to EML files...";
        } else if (e.data == "fileread") {
            document.getElementById("status-text").innerText = "Reading OLM file...";
        }

        if (e.data == "complete") {
            document.getElementById("progress-bar").classList.remove("bg-success");
            document.getElementById("status-text").innerText = "Complete!";
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
        download_btn.click();
    }
}