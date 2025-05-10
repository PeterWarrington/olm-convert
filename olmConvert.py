import sys
import argparse
import zipfile
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from datetime import datetime
import base64
import warnings
import html
import os
import uuid

# Command line interface
def main():
    parser = argparse.ArgumentParser(description='OLM Convert v2.0 - (https://lilpete.me/olm-convert)')

    parser.add_argument('olmPath', type=argparse.FileType('rb'),
                    help='Path to OLM file')
    parser.add_argument('outputDir', type=Path,
                    help='Output directory')

    parser.add_argument('--noAttachments', action='store_true', default=False,
                    help='Don\'t include attachments in output (decreasing file size)')

    parser.add_argument('--format', choices=['eml', 'html'], default="eml",
                    help='Specifies output format')

    parser.add_argument('--verbose', action='store_true', default=False,
                    help='Verbose output')

    args = parser.parse_args()

    print(f"Beginning conversion of {args.olmPath.name}....")
    convertOLM(args.olmPath, args.outputDir, noAttachments=args.noAttachments, format=args.format, verbose=args.verbose)
    print(f"Conversion complete! Files have been written to directory at {args.outputDir}.")

# Convert OLM file specified by olmPath, creating a directory of EML files at outputDir
def convertOLM(olmPath, outputDir, noAttachments=False, verbose=False, format="eml"):
    if format not in ["eml", "html"]:
        raise ValueError(f"Unknown format '{format}', Valid formats are 'eml', 'html'.")

    # OLM file is basically just a zip file, let's access it
    with zipfile.ZipFile(olmPath) as olmZip:
        fileList = olmZip.namelist()

        # Filter file list to just messages using regex
        filterToMessagesRegex = re.compile(r"message_\d{5}.xml$")
        messageFileList = list(filter(filterToMessagesRegex.search, fileList))

        messageListLen = len(messageFileList)
        for messageIndex, messagePath in enumerate(messageFileList):
            # Open xml message file
            messageFile = olmZip.open(messagePath)
            messageFileStr = messageFile.read()

            try:
                # Try to convert message
                messageEml = processMessage(messageFileStr, olmZip=olmZip, noAttachments=noAttachments, format=format)
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
                newPath = os.path.abspath(unresolvedFilePath)[:250] + f".{format}"
                newPath = newPath.replace("\\\\", "/").replace("\\","/")

                # Write converted message to file
                outputFile = open(newPath, "w", encoding='utf-8')
                outputFile.write(messageEmlStr)
                outputFile.close()
                if (verbose):
                    sys.stdout.write(f"[{messageIndex}/{messageListLen}]: Written {newPath}\n")
            except Exception as e:
                print(f"Failed to convert email with subject '{messageEml.subject}' due to error, ignoring. Error detail:")
                print(e)
    if (verbose):
        print("Completed.")

class ConvertedMessage:
    emlString = ""
    subject = ""
    date = ""

    def __init__(self, emlString, subject, date):
        self.emlString = emlString
        self.subject = subject
        self.date = date

# Reads OLM format XML message (xmlString) and returns a ConvertedMessage object containing the message converted to a EML format string
def processMessage(xmlString, olmZip=None, noAttachments=False, format="eml"):
    outputString = ""

    if format == "eml":
        outputString = ""
    elif format in ["html"]:
        try:
            with open("format.html") as f:
                outputString = f.read()
        except:
            raise ValueError(f"Cannot convert to {format}, format.html file not found. Try converting to eml instead.")
    else:
        raise ValueError(f"Unknown format '{format}', Valid formats are 'eml', 'pdf', 'html'.")

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
        dateEmlValue = srcDate.strftime("%a, %d %b %Y %H:%M:%S +0000")
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

    # Set other properties, as per MIME standard (necessary for attachments) 
    mimeVersionEmlStr = "Mime-version: 1.0\n"
    rootBoundaryUUID = generateBoundaryUUID()
    contentTypeEmlStr = f"""Content-type: multipart/mixed;\n\tboundary="{rootBoundaryUUID}"\n\n\n"""

    # Set content boundaries
    contentBoundaryUUID = generateBoundaryUUID()
    multipartAlternativeContentType = f"""Content-type: multipart/alternative;\n\tboundary="{contentBoundaryUUID}"\n\n"""

    # Read plain text content
    plainTextContentTypeEmlStr = """Content-type: text/plain;\n\tcharset="UTF-8"\n"""
    plainTextContentTransferEncodingEmlStr = "Content-transfer-encoding: quoted-printable\n"
    plainTextContentElm = root[0].find("OPFMessageCopyBody")
    if (plainTextContentElm == None or plainTextContentElm.text == None):
       raise ValueError("Text content could not be found in source.")
    plainTextContentStr = lineWrapBody(plainTextContentElm.text)

    # Read HTML content
    HTMLcontentTypeEmlStr = """Content-type: text/html;\n\tcharset="UTF-8"\n"""
    HTMLcontentTransferEncodingEmlStr = "Content-transfer-encoding: quoted-printable\n"
    htmlContentElm = root[0].find("OPFMessageCopyHTMLBody")
    if (htmlContentElm == None or htmlContentElm.text == None):
        warnings.warn("HTML content could not be found in source, ignoring")
    htmlContentRawSrcStr = htmlContentElm.text
    htmlContentEmlStr = html.unescape(htmlContentRawSrcStr) + "\n"
    htmlContentEmlStr = lineWrapBody(htmlContentEmlStr)

    # Read attachments
    attachmentStr = ""
    attachmentList = []
    if (not noAttachments and (olmZip != None)):
        attachmentStr += "\n"
        attachmentListElm = root[0].find("OPFMessageCopyAttachmentList")
        if (attachmentListElm != None):
            attachmentElms = attachmentListElm.findall("messageAttachment")
            for attachmentElm in attachmentElms:
                if format == "eml":
                    attachmentStr += f"--{rootBoundaryUUID}\n"
                    attachmentStr += processAttachment(attachmentElm, olmZip)
                else:
                    attachmentList.append(processAttachment(attachmentElm, olmZip, format=format))

    # Assemble EML message
    if format == "eml":
        outputString += dateEmlStr
        outputString += subjectEmlStr
        outputString += senderEmlStr
        outputString += recipientEmlStr
        outputString += messageIdEmlStr
        outputString += topicEmlStr
        outputString += threadIndexEmlStr
        outputString += mimeVersionEmlStr
        outputString += contentTypeEmlStr
        outputString += f"--{rootBoundaryUUID}\n"
        outputString += multipartAlternativeContentType
        outputString += f"--{contentBoundaryUUID}\n"
        outputString += plainTextContentTypeEmlStr
        outputString += plainTextContentTransferEncodingEmlStr
        outputString += "\n" + plainTextContentStr
        outputString += f"\n--{contentBoundaryUUID}\n"
        outputString += HTMLcontentTypeEmlStr
        outputString += HTMLcontentTransferEncodingEmlStr
        outputString += "\n" + htmlContentEmlStr
        outputString += f"\n--{contentBoundaryUUID}--\n"
        outputString += attachmentStr
        outputString += f"\n--{rootBoundaryUUID}--\n"
    elif format in ["html", "pdf"]:
        outputString = (outputString.replace("$OLM_TITLE", f"{subjectSrcStr} - {sourceDateStr}")
            .replace("$OLM_SUBJECT", subjectSrcStr)
            .replace("$OLM_DATE", dateEmlValue)
            .replace("$OLM_FROM_NAME", senderDetailElm.get("OPFContactEmailAddressName"))
            .replace("$OLM_FROM_EMAIL", senderDetailElm.get("OPFContactEmailAddressAddress"))
            .replace("$OLM_TO_NAME", recipientDetailElm.get("OPFContactEmailAddressName"))
            .replace("$OLM_TO_EMAIL", recipientDetailElm.get("OPFContactEmailAddressAddress"))
            .replace("$OLM_BODY", htmlContentRawSrcStr))

        attachmentsHtmlStr = ""
        if len(attachmentList) > 0:
            for filename, cid, file_url in attachmentList:
                attachmentsHtmlStr += f'<li><a href="{file_url}">{filename}</a></li>'
                outputString = outputString.replace(f"cid:{cid}", file_url)
        else:
            attachmentsHtmlStr = "None"

        outputString = outputString.replace("$OLM_ATTACHMENTS", attachmentsHtmlStr)


    # Return data
    return ConvertedMessage(outputString, subjectSrcStr, dateEmlValue)

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

# Generate a UUID for a MIME boundary, as required by RFC2046 https://datatracker.ietf.org/doc/html/rfc2046#section-5.1.1
def generateBoundaryUUID():
    return "B_" + str(uuid.uuid4().hex)

# Wrap lines of given body to maximum of 78 characters as recommended by RFC 2822 (https://datatracker.ietf.org/doc/html/rfc2822#section-2.1.1)
def lineWrapBody(body):
    return '=\n'.join(body[i:i+77] for i in range(0, len(body), 77)) # This line adapts code from https://stackoverflow.com/a/3258612 (CC BY-SA 2.5)

# Generates EML section (without MIME boundaries) specifying attachments.
def processAttachment(attachmentElm, olmZip, format="eml"):
    # Read content type of attachment
    attachmentContentTypeRaw = attachmentElm.get("OPFAttachmentContentType")
    if (attachmentContentTypeRaw == None):
        raise ValueError("Attachment has no content type.")

    # Read file name of attachment
    attachmentNameRaw = attachmentElm.get("OPFAttachmentName")
    if (attachmentNameRaw == None):
        raise ValueError("Attachment has no name.")

    # Set content type header
    attachmentContentTypeEmlStr = f"""Content-type: {attachmentContentTypeRaw}; name="{attachmentNameRaw}"\n"""

    # Read content ID of attachment (used for example to reference an attached image in html)
    attachmentContentIDraw = attachmentElm.get("OPFAttachmentContentID")
    if (attachmentContentIDraw == None):
        attachmentContentIDEmlStr = ""
    else:
        attachmentContentIDEmlStr = f"Content-ID: <{attachmentContentIDraw}>\n"

    # Set content disposition header
    attachmentContentDispositionEmlStr = f"""Content-disposition: attachment;\n\tfilename="{attachmentNameRaw}"\n"""

    # Set attachment transfer encoding header
    attachmentContentEncodingEmlStr = "Content-transfer-encoding: base64\n"

    # Read attachment file from zip and encode as base64
    attachmentFileZipPath = attachmentElm.get("OPFAttachmentURL")
    if (attachmentFileZipPath == None):
        raise ValueError("Attachment has no specified zip file path.")
    
    try:
        attachmentFileHandle = olmZip.open(attachmentFileZipPath)
        attachmentFileBytes = attachmentFileHandle.read()
        attachmentFileBase64 = base64.b64encode(attachmentFileBytes).decode("utf-8", "ignore")

        # Adds new line to base64 string every 76 characters as base64 strings are decoded as sets
        # of 4 characters therefore the length of each base64 line must be a multiple of 4, and we 
        # want this to be as close to the 78 character RFC2046 recommended  maximum line length as
        # possible.
        attachmentFileBase64 = '\n'.join(attachmentFileBase64[i:i+76] for i in range(0, len(attachmentFileBase64), 76)) # This line adapts code from https://stackoverflow.com/a/3258612 (CC BY-SA 2.5).
        
        attachmentFileHandle.close()
    except Exception as e:
        print(e)
        raise ValueError("Unable to read attachment from OLM zip.")
    
    # Assemble attachment eml
    if format == "eml":
        attachmentResult = ""
        attachmentResult += attachmentContentTypeEmlStr
        attachmentResult += attachmentContentIDEmlStr
        attachmentResult += attachmentContentDispositionEmlStr
        attachmentResult += attachmentContentEncodingEmlStr
        attachmentResult += "\n" + attachmentFileBase64 + "\n"
    else:
        attachmentResult = (attachmentNameRaw, attachmentContentIDraw, f"data:{attachmentContentTypeRaw};base64,{attachmentFileBase64}")

    return attachmentResult

if __name__ == "__main__":
    main()