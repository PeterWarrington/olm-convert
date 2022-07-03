import zipfile
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from datetime import datetime
import base64
import warnings
import html

# Convert OLM file specified by olmPath, creating a directory of EML files at outputDir
def convertOLM(olmPath, outputDir):
    # OLM file is basically just a zip file, let's access it
    with zipfile.ZipFile(olmPath) as olmZip:
        fileList = olmZip.namelist()

        # Filter file list to just messages using regex
        filterToMessagesRegex = re.compile(r"message_\d{5}.xml$")
        messageFileList = list(filter(filterToMessagesRegex.search, fileList))

        # Create output directory
        Path(outputDir).mkdir(parents=True, exist_ok=True)

        i = 0
        for messagePath in messageFileList:
            # Open xml message file
            messageFile = olmZip.open(messagePath)
            messageFileStr = messageFile.read()

            try:
                # Try to convert message
                messageEmlStr = processMessage(messageFileStr)
            except ValueError as e:
                # Couldn't convert this message, skip ip
                print(f"Could not read {messagePath} due to {type(e)}, skipping.")
                print(e)
                continue
            
            # Close messageFile
            messageFile.close()

            # Write converted message to file
            # TODO: Files are named using a simple counter - Should change this to something more meaningful, should create a directory hierarchy 
            outputFile = open(f"./eml_output/{i}.eml", "w")
            outputFile.write(messageEmlStr)
            outputFile.close()

            # Increment counter used to name files
            i += 1

# Reads OLM format XML message (xmlString) and converts it to an EML format string
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

    return emlOutputString

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