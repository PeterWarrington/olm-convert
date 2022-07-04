import sys
import zipfile
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from datetime import datetime
import base64
import warnings
import html
import os

# Command line interface
def main():
    args = sys.argv[1:]

    print("OLM Convert - (https://github.com/PeterWarrington/olm-convert)")
    
    if (len(args) != 2):
        print("USAGE: python3 olmConvert.py <path to OLM file> <output directory>")
    else:
        print(f"Beginning conversion of {args[0]}....")
        convertOLM(args[0], args[1])
        print(f"Conversion complete! Files have been written to directory at {args[1]}.")

# Convert OLM file specified by olmPath, creating a directory of EML files at outputDir
def convertOLM(olmPath, outputDir):
    # OLM file is basically just a zip file, let's access it
    with zipfile.ZipFile(olmPath) as olmZip:
        fileList = olmZip.namelist()

        # Filter file list to just messages using regex
        filterToMessagesRegex = re.compile(r"message_\d{5}.xml$")
        messageFileList = list(filter(filterToMessagesRegex.search, fileList))

        for messagePath in messageFileList:
            # Open xml message file
            messageFile = olmZip.open(messagePath)
            messageFileStr = messageFile.read()

            try:
                # Try to convert message
                messageEml = processMessage(messageFileStr)
                messageEmlStr = messageEml.emlString
            except ValueError as e:
                # Couldn't convert this message, skip ip
                print(f"Could not read {messagePath} due to {type(e)}, skipping.")
                print(e)
                continue
            
            # Close messageFile
            messageFile.close()

            # Assemble path to message
            accountsPathIndex = messagePath.find("Accounts/") + 9
            accountName = messagePath[accountsPathIndex:messagePath.find("/", accountsPathIndex)]

            emailFolderIndex = messagePath.find("com.microsoft.__Messages/") + 25
            emailFolder = messagePath[emailFolderIndex:messagePath.rfind("/")]

            emailPathStr = f"{outputDir}/{accountName}/{emailFolder}"
            Path(emailPathStr).mkdir(parents=True, exist_ok=True)

            # Assemble safe eml file name
            try:
                fileName = f"{messageEml.subject[:200]} - {messageEml.date}".replace(":", ".").replace("/", ".").replace("\\", ".")
                fileName = ''.join([c for c in fileName if c.isalpha() or c.isdigit() or c==' ' or c=='.']).rstrip()
                unresolvedFilePath = f"{emailPathStr}/{fileName}"
                newPath = os.path.abspath(unresolvedFilePath)[:250] + ".eml"
                newPath = newPath.replace("\\\\", "/").replace("\\","/")

                # Write converted message to file
                outputFile = open(newPath, "w", encoding='utf-8')
                outputFile.write(messageEmlStr)
                outputFile.close()
            except Exception as e:
                print(f"Failed to convert email with subject '{messageEml.subject}' due to error, ignoring. Error detail:")
                print(e)

class ConvertedMessage:
    emlString = ""
    subject = ""
    date = ""

    def __init__(self, emlString, subject, date):
        self.emlString = emlString
        self.subject = subject
        self.date = date

# Reads OLM format XML message (xmlString) and returns a ConvertedMessage object containing the message converted to a EML format string
def processMessage(xmlString):
    emlOutputString = ""

    # Read from file
    root = ET.fromstring(xmlString)

    # Check number of <email> elements below <emails> root element is 1
    if (len(root) != 1):
        raise ValueError(f"Expected one email beneath root, got {len(root)}.")
    
    # Read date of email (e.g. "2022-05-26T17:55:40" -> "Thu, 26 May 2022 18:55:40 +0100") (https://datatracker.ietf.org/doc/html/rfc5322#section-3.3)
    sourceDateElm = root[0].find("OPFMessageCopySentTime")
    if (sourceDateElm == None):
        raise ValueError("Sent date could not be found in source.")

    sourceDateStr = sourceDateElm.text
    try:
        srcDate = datetime.strptime(sourceDateStr, "%Y-%m-%dT%X")
        dateEmlValue = srcDate.strftime("%a, %m %B %Y %X %z")
        dateEmlStr = f"Date: {dateEmlValue}\n"
    except ValueError:
        raise ValueError("Unexpected sent time format in source.")
    
    # Read subject of email
    subjectElm = root[0].find("OPFMessageCopySubject")
    if (subjectElm == None):
        # Subject could not be found in source, just set the subject to be empty
        subjectSrcStr = ""
        subjectEmlStr = f"Subject: \n"
    else:
        subjectSrcStr = subjectElm.text 

        subjectEmlValue = headerEncode(subjectSrcStr)
        subjectEmlStr = f"Subject: {subjectEmlValue}\n"

    # Read From field of email
    senderElm = root[0].find("OPFMessageCopySenderAddress")
    if (senderElm == None):
        # Sender address could not be found in source, just don't set a sender address in EML header
        senderEmlStr = ""
    else:
        if (len(senderElm) != 1):
            raise ValueError(f"Expected 1 child of sender address, got {len(senderElm)}")
        
        senderDetailElm = senderElm.find("emailAddress")
        if (senderDetailElm == None):
            raise ValueError("No child email address detail element for sender found.")

        senderEmlValue = addressEncode(senderDetailElm)
        senderEmlStr = f"From: {senderEmlValue}\n"

    # Read To field of email (does not support multiple people copied into email, in that case chooses 1st person copied to)
    recipientElm = root[0].find("OPFMessageCopyToAddresses")
    if (recipientElm == None):
        # Recipient address could not be found in source, just don't set a recipient address in EML header
        recipientEmlStr = ""
    else:
        if (len(recipientElm) != 1):
            warnings.warn(f"Multiple recipient addresses, choosing first only.")
        
        recipientDetailElm = recipientElm.find("emailAddress")
        if (recipientDetailElm == None):
            raise ValueError("No child email address detail element for recipient found.")

        recipientEmlValue = addressEncode(recipientDetailElm)
        recipientEmlStr = f"To: {recipientEmlValue}\n"

    # Read message ID of email
    messageIdElm = root[0].find("OPFMessageCopyMessageID")
    if (messageIdElm == None):
        raise ValueError("Message ID could not be found in source.")

    messageIdEmlValue = messageIdElm.text
    messageIdEmlStr = f"Message-ID: <{messageIdEmlValue}>\n"

    # Read thread topic of email
    topicElm = root[0].find("OPFMessageCopyThreadTopic")
    if (topicElm == None or topicElm.text == None):
        # Thread topic could not be found in source, just don't set a thread topic in EML header
        topicEmlStr = ""
    else:
        topicSrcStr = topicElm.text 
        topicEmlValue = headerEncode(topicSrcStr)
        topicEmlStr = f"Thread-Topic: {topicEmlValue}\n"

    # Read thread index of email
    threadIndexElm = root[0].find("OPFMessageCopyThreadIndex")
    if (threadIndexElm == None):
        raise ValueError("Thread index could not be found in source.")
    
    threadIndexEmlValue = threadIndexElm.text
    threadIndexEmlStr = f"Thread-Index: {threadIndexEmlValue}\n"

    # Set other properties
    mimeVersionEmlStr = "Mime-version: 1.0\n"
    contentTypeEmlStr = """Content-type: text/html;\n\tcharset="UTF-8"\n"""
    contentTransferEncodingEmlStr = "Content-transfer-encoding: quoted-printable\n"

    # Read HTML content
    htmlContentElm = root[0].find("OPFMessageCopyHTMLBody")
    if (htmlContentElm == None or htmlContentElm.text == None):
        raise ValueError("HTML content could not be found in source.")
    
    htmlContentRawSrcStr = htmlContentElm.text
    htmlContentEmlStr = html.unescape(htmlContentRawSrcStr)

    # Assemble EML message
    emlOutputString += dateEmlStr
    emlOutputString += subjectEmlStr
    emlOutputString += senderEmlStr
    emlOutputString += recipientEmlStr
    emlOutputString += messageIdEmlStr
    emlOutputString += topicEmlStr
    emlOutputString += threadIndexEmlStr
    emlOutputString += mimeVersionEmlStr
    emlOutputString += contentTypeEmlStr
    emlOutputString += contentTransferEncodingEmlStr
    emlOutputString += "\n" + htmlContentEmlStr

    # Return data
    return ConvertedMessage(emlOutputString, subjectSrcStr, dateEmlValue)

# Convert email header value to RFC2047 base64 encoded UTF-8 string (https://datatracker.ietf.org/doc/html/rfc2047)
def headerEncode(value):
    B64Str = base64.b64encode(value.encode("utf-8", "ignore")).decode("utf-8", "ignore")
    return f"=?UTF-8?B?{B64Str}?="

# Convert <emailAddress> element to a email header value
def addressEncode(addressElm):
    rawName = addressElm.get("OPFContactEmailAddressName")
    if (rawName == None):
        warnings.warn("Email address name not found, ignoring.")
        rawName = ""
    encodedName = headerEncode(rawName)

    address = addressElm.get("OPFContactEmailAddressAddress")
    if (address == None):
        raise ValueError("Email address not found in emailAddress element.")

    return f"{encodedName} <{address}>"

if __name__ == "__main__":
    main()