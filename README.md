# jotlmsg
It's a simple API meant to easily generate Microsoft Outlook message files (.msg). 
This library is based on [Apache POI](https://poi.apache.org) and is currenly a work in progress groovy impementation of [ctabin's jotlmsg](https://github.com/ctabin/jotlmsg)

## Installation

Since this isn't production ready yet (still failing unit tests) i do not encourage ANYBODY to use this. This is, currently, a personal project i'm working on, and thus, not ready.

## Usage examples

Create a new message:
```Java
OutlookMessage message = new OutlookMessage();
message.setSubject("Hello");
message.setPlainTextBody("This is a message draft.");

//creates a new Outlook Message file
message.writeTo(new File("myMessage.msg"));

//creates a javax.mail MimeMessage
MimeMessage mimeMessage = message.toMimeMessage();
```

Read an existing message:
```Java
OutlookMessage message = new OutlookMessage(new File("aMessage.msg"));
System.out.println(message.getSubject());
System.out.println(message.getPlainTextBody());
```

Managing recipients:
```Java
OutlookMessage message = new OutlookMessage();
message.addRecipient(Type.TO, "cedric@jotlmsg.com");
message.addRecipient(Type.TO, "bill@microsoft.com", "Bill");
message.addRecipient(Type.CC, "steve@apple.com", "Steve");
message.addRecipient(Type.BCC, "john@gnu.com");
        
List<OutlookMessageRecipient> toRecipients = message.getRecipients(Type.TO);
List<OutlookMessageRecipient> ccRecipients = message.getRecipients(Type.CC);
List<OutlookMessageRecipient> bccRecipients = message.getRecipients(Type.BCC);
List<OutlookMessageRecipient> allRecipients = message.getAllRecipients();
```

Managing attachments:
```Java
OutlookMessage message = new OutlookMessage();
message.addAttachment("aFile.txt", "text/plain", new FileInputStream("data.txt")); //will be stored in memory
message.addAttachment("aDocument.pdf", "application/pdf", new FileInputStream("file.pdf")); //will be stored in memory
message.addAttachment("hugeFile.zip", "application/zip", a -> new FileInputStream("data.zip")); //piped to output stream

List<OutlookMessageAttachment> attachments = message.getAttachments();
```

## Limitations

The current implementation allows to create simple msg files with many recipients (up to 2048) and attachments (up to 2048). 
However, there is not current support of Microsoft Outlook advanced features like appointments or calendar integration, nor embedded messages.
