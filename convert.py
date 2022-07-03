import xml.etree.ElementTree as ET
from datetime import datetime
import base64
import warnings
import html

# Reads OLM format XML message at inputFilePath and converts it to an EML file at outputFilePath
def processMessage(inputFilePath, outputFilePath):
    emlOutputString = ""

    # Read from file
    xmlTree = ET.parse(inputFilePath)
    root = xmlTree.getroot()

    # Check number of <email> elements below <emails> root element is 1
    if (len(root) != 1):
        raise ValueError(f"Expected one email beneath root, got {len(root)}.")
    
    # Read date of email (e.g. "2022-05-26T17:55:40" -> "Thu, 26 May 2022 18:55:40 +0100")
    sourceDateElm = xmlTree.getroot()[0].find("OPFMessageCopySentTime")
    if (sourceDateElm == None):
        raise ValueError("Sent date could not be found in source.")

    sourceDateStr = sourceDateElm.text
    try:
        srcDate = datetime.strptime(sourceDateStr, "%Y-%m-%dT%X")
        dateEmlValue = srcDate.strftime("%a, %m %B %Y %X %z")
        dateEmlStr = f"Date: {dateEmlValue}"
    except ValueError:
        raise ValueError("Unexpected sent time format in source.")
    
    # Read subject of email
    subjectElm = xmlTree.getroot()[0].find("OPFMessageCopySubject")
    if (subjectElm == None):
        raise ValueError("Subject could not be found in source.")
    
    subjectSrcStr = subjectElm.text 

    subjectEmlValue = headerEncode(subjectSrcStr)
    subjectEmlStr = f"Subject: {subjectEmlValue}"

    # Read From field of email
    senderElm = xmlTree.getroot()[0].find("OPFMessageCopySenderAddress")
    if (senderElm == None):
        raise ValueError("Sender address could not be found in source.")
    if (len(senderElm) != 1):
        raise ValueError(f"Expected 1 child of sender address, got {len(senderElm)}")
    
    senderDetailElm = senderElm.find("emailAddress")
    if (senderDetailElm == None):
        raise ValueError("No child email address detail element for sender found.")

    senderEmlValue = addressEncode(senderDetailElm)
    senderEmlStr = f"From: {senderEmlValue}"

    # Read To field of email (does not support multiple people copied into email, in that case chooses 1st person copied to)
    recipientElm = xmlTree.getroot()[0].find("OPFMessageCopyToAddresses")
    if (recipientElm == None):
        raise ValueError("Recipient address could not be found in source.")
    if (len(recipientElm) != 1):
        warnings.warn(f"Multiple recipient addresses, choosing first only.")
    
    recipientDetailElm = recipientElm.find("emailAddress")
    if (recipientDetailElm == None):
        raise ValueError("No child email address detail element for recipient found.")

    recipientEmlValue = addressEncode(recipientDetailElm)
    recipientEmlStr = f"To: {recipientEmlValue}"

    # Read message ID of email
    messageIdElm = xmlTree.getroot()[0].find("OPFMessageCopyMessageID")
    if (messageIdElm == None):
        raise ValueError("Message ID could not be found in source.")

    messageIdEmlValue = messageIdElm.text
    messageIdEmlStr = f"Message-ID: <{messageIdEmlValue}>"

    # Read thread topic of email
    topicElm = xmlTree.getroot()[0].find("OPFMessageCopyThreadTopic")
    if (topicElm == None):
        raise ValueError("Thread topic could not be found in source.")
    
    topicSrcStr = topicElm.text 

    topicEmlValue = headerEncode(topicSrcStr)
    topicEmlStr = f"Thread-Topic: {topicEmlValue}"

    # Read thread index of email
    threadIndexElm = xmlTree.getroot()[0].find("OPFMessageCopyThreadIndex")
    if (threadIndexElm == None):
        raise ValueError("Thread index could not be found in source.")
    
    threadIndexEmlValue = threadIndexElm.text
    threadIndexEmlStr = f"Thread-Index: {threadIndexEmlValue}"

    # Set other properties
    mimeVersionEmlStr = "Mime-version: 1.0"
    contentTypeEmlStr = """Content-type: text/html;\n\tcharset="UTF-8" """
    contentTransferEncodingEmlStr = "Content-transfer-encoding: quoted-printable"

    # Read HTML content
    htmlContentElm = xmlTree.getroot()[0].find("OPFMessageCopyHTMLBody")
    if (htmlContentElm == None):
        raise ValueError("HTML content could not be found in source.")
    
    htmlContentRawSrcStr = htmlContentElm.text
    htmlContentEmlStr = html.unescape(htmlContentRawSrcStr)

    # Assemble EML message
    emlOutputString += dateEmlStr + "\n"
    emlOutputString += subjectEmlStr + "\n" 
    emlOutputString += senderEmlStr + "\n"
    emlOutputString += recipientEmlStr + "\n"
    emlOutputString += messageIdEmlStr + "\n"
    emlOutputString += topicEmlStr + "\n"
    emlOutputString += threadIndexEmlStr + "\n"
    emlOutputString += mimeVersionEmlStr + "\n"
    emlOutputString += contentTypeEmlStr + "\n"
    emlOutputString += contentTransferEncodingEmlStr + "\n"
    emlOutputString += "\n" + htmlContentEmlStr

    # Write eml message
    writeHandle = open(outputFilePath, "w")
    writeHandle.write(emlOutputString)
    writeHandle.close()

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