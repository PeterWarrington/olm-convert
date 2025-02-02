import sys
from pyodide.http import pyfetch
import re
import shutil
from io import BytesIO
import zipfile
import os

async def init():
    response = await pyfetch("https://raw.githubusercontent.com/PeterWarrington/olm-convert/refs/tags/v1.2.0/olmConvert.py")
    olm_convert_text = (await response.bytes()).decode("utf-8")
    with open("olmConvert.py", "w") as f:
        f.write(olm_convert_text)

async def convert():
    from js import postMessage
    from js import olmFileBytes
    from js import includeAttachments
    import olmConvert

    try:
        postMessage("fileread")
        olmFileBytes = BytesIO(olmFileBytes.to_py())
        print(type(olmFileBytes))
        
        class FakeStdout():
            def __init__(self, old_stdout):
                self.old_stdout = old_stdout

            def write(self, string):
                progress_match = re.search("^\\[(\\d+)\\/(\\d+)]\\:", string)
                if (progress_match):
                    index = int(progress_match.group(1))
                    out_of = int(progress_match.group(2))
                    progressPercent = (index / out_of) * 100
                    postMessage("progress:{:.2f}%".format(round(progressPercent, 2)))
                else:
                    if self.old_stdout is not None:
                        self.old_stdout.write(string)

            def flush(self):
                pass

        class FakeStderr():
            def __init__(self, old_stderr):
                self.old_stderr = old_stderr

            def write(self, string):
                if self.old_stderr is not None:
                    self.old_stderr.write(string)

            def flush(self):
                pass

        sys.stdout = FakeStdout(sys.stdout)
        sys.stderr = FakeStderr(sys.stderr)

        postMessage("startconvert")
        output_dir = "/olm-convert-output"
        olmConvert.convertOLM(olmFileBytes, output_dir, noAttachments=(not includeAttachments), verbose=True)
        postMessage("progress:100%")

        postMessage("createzip")
        with zipfile.ZipFile('/outputEmls.zip', 'w', zipfile.ZIP_DEFLATED) as zip:
            out_of = sum(len(files) for _, _, files in os.walk(output_dir))
            fileCount = 0
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    progressPercent = (fileCount / out_of) * 100
                    postMessage("progress:{:.2f}%".format(round(progressPercent, 2)))
                    zip.write(os.path.join(root, file), 
                            os.path.relpath(os.path.join(root, file), 
                                            os.path.join(output_dir, '..')))
                    
                    # Remove files in temp conversion dir as we go to save RAM
                    # (file system is stored in RAM on web)
                    os.remove(os.path.join(root, file)) 

                    fileCount += 1
        postMessage("progress:0%")
    except Exception as e:
        postMessage(f"error:{str(e)}")