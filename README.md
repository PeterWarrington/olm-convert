# OLM Convert

A python utility to convert an Outlook for Mac archive (.olm file), that can only be opened with Outlook for Mac, to a set of standard .eml files that can be opened by almost all email clients.

Output EML messages are organised hierarchically, for example a message contained in an OLM file will be output to an EML file like `"<output directory>/example@example.com/Inbox/Subject - Mon, 04 July 2022 21.04.56.eml"`.

Can be used as a command line interface or as a module.

Does not yet support attachments and can only output to an EML file with HTML content type.

## Usage

Command line:
```
python3 olmConvert.py <path to OLM file> <output directory>
```

## Module reference


### convertOLM(olmPath, outputDir)
Convert OLM file specified by olmPath, creating a directory of EML files at outputDir.

### processMessage(xmlString)
Reads OLM format XML message (xmlString) and returns a ConvertedMessage object containing the message converted to a EML format string. Can also return ValueError.

### headerEncode(value)
Converts email header string value to RFC2047 base64 encoded UTF-8 string (<https://datatracker.ietf.org/doc/html/rfc2047>).

### addressEncode(addressElm)
Converts &lt;emailAddress&gt; element ([xml.etree.ElementTree.Element](https://docs.python.org/3/library/xml.etree.elementtree.html#xml.etree.ElementTree.Element)) to an email header value.
