const worker = new Worker("olmConvertWorker.js");

// Download blob code from https://stackoverflow.com/a/33622881
var downloadBlob, downloadURL;
downloadBlob = function(data, fileName, mimeType) {
  
  var blob, url;
  blob = new Blob([data], {
    type: mimeType
  });
  console.log("finished creating blob, creating object url...");
  url = window.URL.createObjectURL(blob);
  console.log("finished creating object url, downloading url...");
  downloadURL(url, fileName);
};

downloadURL = function(data, fileName) {
  var a;
  a = document.createElement('a');
  a.href = data;
  a.download = fileName;
  document.body.appendChild(a);
  a.style = 'display: none';
  a.click();
  a.remove();
};

// End download blob code

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
    if (typeof(e.data) == "string") {
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
            worker.terminate();
            document.getElementById("progress-bar").classList.remove("bg-success");
            document.getElementById("status-text").innerText = "Complete!";
        }
    } else {
        console.log("received data");
        downloadBlob(e.data, "emlOutput.zip", "application/zip");
    }
}