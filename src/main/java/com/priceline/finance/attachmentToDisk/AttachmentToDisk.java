package com.priceline.finance.attachmentToDisk;

import java.io.File;
import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;

public class AttachmentToDisk {

	private static String attachmentsDirectoryPath = "C:\\Attachments\\";
	private static Properties properties;
	private static Session session;

	public static void main(final String[] args) {
		properties = new Properties();
		properties.setProperty("mail.host", "imap.gmail.com");
		properties.setProperty("mail.port", "995");
		properties.setProperty("mail.transport.protocol", "imaps");
		session = Session.getInstance(properties, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(args[0], args[1]);
			}
		});
		try {
			Store store = session.getStore("imaps");
			store.connect();
			Folder inboxFolder = store.getFolder("Personal");
			inboxFolder.open(Folder.READ_ONLY);
			Message[] messages = inboxFolder.getMessages();
			for (int i = 0; i < messages.length; i++) {
				Message message = messages[i];
				if (message.getContent() instanceof Multipart) {
					Multipart multiPart = (Multipart) message.getContent();
					// for (int j = 0; j < multiPart.getCount(); j++) {
					for (int j = 0; j < 10; j++) {
						MimeBodyPart bodyPart = (MimeBodyPart) multiPart
								.getBodyPart(j);
						if (bodyPart.getDisposition() != null
								&& bodyPart.getDisposition().equalsIgnoreCase(
										Part.ATTACHMENT)) {
							bodyPart.saveFile(attachmentsDirectoryPath
									+ File.separator + bodyPart.getFileName());
						}
					}
				}
			}
			inboxFolder.close(true);
			store.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
}