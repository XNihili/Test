package test;
import javax.mail.*;
import javax.mail.Message.RecipientType;
import com.sun.mail.imap.IMAPFolder;
import java.util.Properties;
import java.util.ArrayList;

public class MailReader {
	
	final String PROTOCOL = "imap";
	final String HOST = "outlook.office365.com";
	final int PORT = 993;
	//final String PORT2 = "995";
	final boolean DEBUG = false;
	final String MAILACHERCHER1="XNIHILI@HOTMAIL.COM";
	final String MAILACHERCHER2="XNihili@hotmail.com";
	
	final static String LOGIN = "XNihili@hotmail.com";
	final static String PASSWORD = "8&oss@@tv&Fx*D";
	final static String FILE = "e:\\test.xlsx";
	
	/**
	 * private Session getMessage (String username, String password)
	 * Ouvre une connexion autentifiée avec le serveur mail
	 * renvoie le contenu de INBOX ou NULL en cas d'erreur
	 */
	private Message[] getMessage(String username, String password) {
		Session session = this.getImapSession();
		Message[] messages=null;
		try {
			Store store = session.getStore("imap");
			store.connect(HOST, PORT, username, password);
			IMAPFolder inbox = (IMAPFolder)store.getFolder("INBOX");
			inbox.open(Folder.READ_ONLY);
			messages = inbox.getMessages();
		} catch (Exception e) {
			System.out.println("Exception in getting mail : "+e.getMessage());
		}
		return messages;
	}
	
	
	/**
	 * private boolean verificationDestinataire(Address[] toList, Address[] ccList)
	 * vérifie si les destinataires à rechercher sont dans la liste toList et ccList
	 * retourne true s'il y en a un présent, false sinon
	 */
	private boolean verificationDestinataire(Address[] toList, Address[] ccList) {
		try {
			if(toList!=null)
			for(int i=0;i<toList.length;i++) {
				String to = toList[i].toString();
				if(MAILACHERCHER1.equals(to))
					return true;
				if(MAILACHERCHER2.equals(to)) 
					return true;
			}
			if(ccList!=null)
			for(int i=0;i<ccList.length;i++) {
				String cc = ccList[i].toString();
				if(MAILACHERCHER1.equals(cc))
					return true;
				if(MAILACHERCHER2.equals(cc)) 
					return true;
			}
		} catch (Exception e) {
			System.out.println("Exception To or CC field containing data : "+e.getMessage());
		}
		return false;
	}
	
	/**
	 * Message[] triMessage(Message[] message)
	 * Fait le tri des mails à afficher
	 * Renvoi les mails sélectionnés
	 */
	private Message[] triMessage(Message[] messages) {
		 ArrayList<Message> listeMessage = new ArrayList<>();
		 Message[] messageRetour=null;
		try {
			int k=0;
			for (int i = 0; i < messages.length; i++) {
				Message msg = messages[i];
				Address[] toList = msg.getRecipients(RecipientType.TO);
				Address[] ccList = msg.getRecipients(RecipientType.CC);
				try {
					if(verificationDestinataire(toList, ccList) ) {
						listeMessage.add(msg);
						k++;
					}
				} catch (Exception e) {
					System.out.println("Exception analyzing To and CC field : "+e.getMessage());
				}
			}
			messageRetour=new Message[k];
			for(int i=0; i<k;i++) {
				messageRetour[i]=listeMessage.get(i);
			}
		} catch (Exception e) {
			System.out.println("Exception in filtering mail : "+e.getMessage());
		}
		return messageRetour;
	}
	
	/**
	 * private Session getImapSession(){
	 * renvoie une session pour le serveur mail
	 */
	private Session getImapSession(){	
		Properties props = new Properties();
		props.setProperty("mail.store.protocol", PROTOCOL);
		props.setProperty("mail.debug", Boolean.toString(DEBUG));
		props.setProperty("mail.imap.host",HOST);
		props.setProperty("mail.imap.port", Integer.toString(PORT));
		props.setProperty("mail.imap.ssl.enable","true");
		Session session = Session.getDefaultInstance(props, null);
		session.setDebug(DEBUG);
		return session;
	}
	
	/**
	 *  Main
	 */
	public static void main(String[] args) {
		MailReader mailReader = new MailReader();
		ExcelWriter excelWriter = new ExcelWriter();
		Message[] messages = mailReader.triMessage(mailReader.getMessage(LOGIN,PASSWORD));
		try {
			for (int i = 0; i < messages.length; i++) {
				Message msg = messages[i];
				Address[] fromAddress = msg.getFrom();
				String subject = msg.getSubject();
				excelWriter.insertDataInExcelFileXSSF(FILE, "Mail", subject);
			}
		} catch (Exception e) {
			System.out.println("Exception in reading mail : "+e.getMessage());
		}
		WebReader webreader = new WebReader();
		webreader.getContentFromURWithJsoup("https://developer.mozilla.org/en-US/docs/Learn/HTML/Tables/Basics");
	}
}
