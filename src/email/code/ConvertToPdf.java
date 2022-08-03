package email.code;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.FileSpecification;
import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;
import com.chilkatsoft.CkEmail;
import com.chilkatsoft.CkEmailBundle;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkMailboxes;
import com.chilkatsoft.CkMessageSet;
public class ConvertToPdf{
	static CkImap imap;
ConvertToPdf(CkImap imap1,String destination,String filetype,ArrayList<String> pstfolderlist,Date startdate,Date enddate) {
	//public static void convertPdf(CkImap imap,String destination,String filetype,ArrayList<String> pstfolderlist) {
		
		this.imap=imap1;
		File f = null;

		boolean success;
		CkMailboxes mboxes = imap.ListMailboxes("", "*");

		if (imap.get_LastMethodSuccess() == false) {
			System.out.println(imap.lastErrorText());
			return; 
		}
//		textField.setEnabled(false);
//		textField_1.setEnabled(false);
//		lblNewLabel_2.setText("Connected");
		long countdest = 0;
//		File file = new File(
//				System.getProperty("user.home") + "/Desktop" + File.separator + "Test");
		int j = 0;
		int l=0;
//		PersonalStorage pst = PersonalStorage.create(file.getAbsolutePath() + ".pst",
//				FileFormatVersion.Unicode);
		//if(filetype.equalsIgnoreCase("pdf")) {
		while (j < mboxes.get_Count()) {
			if (Main_Frame.stop) {
				break;
			}
			String a=mboxes.getName(j).replaceAll("/", "\\\\");
			a = a.replaceAll("[\\[\\]\\(\\)]", "");
			System.out.println("printing a---"+a);
			
			if(pstfolderlist.contains(a) ) {
			
			
			System.out.println("41-----"+mboxes.getName(j));

			String foldername = mboxes.getName(j).replace("/", File.separator);
			String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
		//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
System.out.println("51--"+destination);
System.out.println("52--"+foldername);
		

//			lblNewLabel_3_1.setText("Folder Name :" + foldername);
			CkEmailBundle bundle;
			CkMessageSet messageSet;
			// We can choose to fetch UIDs or sequence numbers.
			boolean fetchUids = true;
			success = imap.SelectMailbox(mboxes.getName(j));
			if (success != true) {
				System.out.println(imap.lastErrorText());
				// return;
			}
			// Get the message IDs of all the emails in the mailbox
			messageSet = imap.Search("ALL", fetchUids);
			if (imap.get_LastMethodSuccess() == true) {
				if (imap.get_LastMethodSuccess() == false) {
					System.out.println(imap.lastErrorText());
					// return;
				}

			//	bundle = imap.FetchBundle(messageSet);
				if (imap.get_LastMethodSuccess() == false) {
					System.out.println(imap.lastErrorText());
					// return;
				}

				int numEmails = imap.get_NumMessages();
				if(numEmails>0) {
					 f = new File(destination+ File.separator+ foldername);
						f.mkdirs();
					}
				 boolean bUid = false;
				int i = 1;
				System.out.println(numEmails);
				Main_Frame.listduplicacy.clear();
				while (i <= numEmails) {
					boolean mail_convert = false;
					try {
						if (Main_Frame.demo) {
							if (i-1 == All_Data.demo_count) {
								break;
							}
						}
						if (Main_Frame.stop) {
							break;
						}
						
						while(Main_Frame.pause) {
							Main_Frame.lbl_progressreport.setText("<html><b>process has been paused");
						}
						
						System.out.println("printing i : "+i);
						CkEmail email = imap.FetchSingle(i,bUid);
						if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
							//	message.getAttachments().clear();
								System.out.println("selected..");
							email.DropAttachments();
							}
						email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
						MailMessage message = MailMessage
								.load(System.getProperty("user.home") + File.separator + "0.eml");
						String path5 = f.getAbsolutePath() + File.separator + i
								+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
						System.out.println("date--"+message.getDate());
						Date reciveddate=message.getDate();
					//	DateFormat istFormat = new SimpleDateFormat();
						//  DateFormat gmtFormat = new SimpleDateFormat();
						//  TimeZone gmtTime = TimeZone.getTimeZone("GMT");
						//  TimeZone istTime = TimeZone.getTimeZone("IST");
						  reciveddate.toGMTString();
						  //istFormat.setTimeZone(gmtTime);
//						  gmtFormat.setTimeZone(istTime);
						
						//reciveddate.set
						// utcFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
						//imap.del
				
				message.setDate(reciveddate);
				System.out.println("date2--"+message.getDate());
						if(Main_Frame.chckbxRemoveDuplicacy.isSelected()) {


							String input = Main_Frame.duplicacymail(message);

							if (!Main_Frame.listduplicacy.contains(input)) {

								Main_Frame.listduplicacy.add(input);

								if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
									if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
										mail_convert=	MailConvert(path5, message, filetype, email,f);
										Main_Frame.count_destination++;
									} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
										mail_convert=	MailConvert(path5, message, filetype, email,f);
										Main_Frame.count_destination++;
									}
								}else {
									mail_convert=	MailConvert(path5, message, filetype, email,f);
									Main_Frame.count_destination++;
								}
							} else {
								System.out.println(" duplicate message");
								System.out.println(input);

							}
						
						}else {
							
						
						if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
							if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
								mail_convert=	MailConvert(path5, message, filetype, email,f);
								Main_Frame.count_destination++;
							} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
								mail_convert=MailConvert(path5, message, filetype, email,f);
								Main_Frame.count_destination++;
							}
						}else {
							mail_convert=MailConvert(path5, message, filetype, email,f);
							Main_Frame.count_destination++;
						}
						}	
					//	MailConvert(path5, message, filetype, email,f);
						//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
						
//						ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//						message.save(emlStream, SaveOptions.getDefaultMhtml());
//						LoadOptions lo = new LoadOptions();
//						lo.setLoadFormat(LoadFormat.MHTML);
//						Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//						path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//						doc.save(path5, SaveFormat.PDF);
//						//email.
////						If email.NumAttachedMessages > 0 then
////				        '''how do I get the original email and download the file attachments if any.
////				    else
////				        Dim i as integer
////				        For i = 0 to email.NumAttachments - 1
////				            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////				        next i
////						PdfContentEditor editor = new PdfContentEditor();
////						editor.bindPdf(path5);
//						if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								if (Main_Frame.chckbxSavePdfAttachment.isSelected()) {
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								}
//								else {
//								File f1;
//								
//								if (System.getProperty("os.name").toLowerCase().contains("windows")) {
//									 f1=new File(System.getProperty("java.io.tmpdir"));
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
//								}
//								else {
//									f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
//											+ "Application Support" );
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
//								}
//		
//	com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
//
//		// Set up a new file to be added as attachment
//		FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
//
//		// Add an attachment to document's attachment collection
//		pdfDocument.getEmbeddedFiles().add(fileSpecification);
//
//		// Save the updated document
//		pdfDocument.save(path5);
//		File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
//		f1.delete();
//		f2.delete();
//							}
//							}
//						}
						
//						if (message.getAttachments().size() > 0) {
//							
//
////							MapiMessage msg = MapiMessage.fromMailMessage(message);
////							PdfContentEditor editor = new PdfContentEditor();
////							editor.bindPdf(path5);
////
////							for (MapiAttachment attachment : msg.getAttachments()) {
////								if (!(attachment.getExtension() == null)) {
////								ByteArrayOutputStream eml = new ByteArrayOutputStream();
////
////								attachment.save(eml);
////								editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
////										attachment.getDisplayName(), "");
////								}
////							}
//						}
						//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
//						lblNewLabel_3.setText("count " + i);

//						lblNewLabel_3_2.setText("Total Count :" + countdest++);
						Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
								+ fname + "   Extarcting messsage " + message.getSubject());
						
						if(Main_Frame.chckbxDeleteEmailFrom.isSelected()) {
							if(mail_convert) {
						//	imap.SetFlags(messageSet,"Deleted",i);	
							success = imap.SetMailFlag(email,"Deleted",1);
						}
						
					}
					} catch (Exception e) {
						System.out.println("inside the catchg block");
						e.printStackTrace();
					//Main_Frame main_Frame = new Main_Frame(null, 0);
					Main_Frame.connectionHandle();
					System.out.println("going to select mailbox");
					success = imap.SelectMailbox(mboxes.getName(j));
					if (success != true) {
						System.out.println(imap.lastErrorText());
						// return;
					}
						//continue;
					System.out.println("reconnect..."+success);
					i--;
					}
					
						
						
					
				//	imap.   .FetchSingle(i,bUid).d
				//	mboxes.
					i++;
			
				}

			}
		}
			j++;
		}
	//}
//		else  if(filetype.equalsIgnoreCase("doc")) {
//
//			while (j < mboxes.get_Count()) {
//				if (Main_Frame.stop) {
//					break;
//				}
//				String a=mboxes.getName(j).replaceAll("/", "\\\\");
//				a = a.replaceAll("[\\[\\]\\(\\)]", "");
//				System.out.println("printing a---"+a);
//				
//				if(pstfolderlist.contains(a) ) {
//				
//				
//				System.out.println("41-----"+mboxes.getName(j));
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+foldername);
//			
//
////				lblNewLabel_3_1.setText("Folder Name :" + foldername);
//				CkEmailBundle bundle;
//				CkMessageSet messageSet;
//				// We can choose to fetch UIDs or sequence numbers.
//				boolean fetchUids = true;
//				success = imap.SelectMailbox(mboxes.getName(j));
//				if (success != true) {
//					System.out.println(imap.lastErrorText());
//					// return;
//				}
//				// Get the message IDs of all the emails in the mailbox
//				messageSet = imap.Search("ALL", fetchUids);
//				if (imap.get_LastMethodSuccess() == true) {
//					if (imap.get_LastMethodSuccess() == false) {
//						System.out.println(imap.lastErrorText());
//						// return;
//					}
//
//					//	bundle = imap.FetchBundle(messageSet);
//					if (imap.get_LastMethodSuccess() == false) {
//						System.out.println(imap.lastErrorText());
//						// return;
//					}
//
//					int numEmails = imap.get_NumMessages();
//					if(numEmails>0) {
//						 f = new File(destination+ File.separator+ foldername);
//							f.mkdirs();
//						}
//					 boolean bUid = false;
//					int i = 1;
//					System.out.println(numEmails);
//					while (i <= numEmails) {
//						try {
//							if (Main_Frame.stop) {
//								break;
//							}	
//							System.out.println(i);
//							CkEmail email = imap.FetchSingle(i,bUid);
//							if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//								//	message.getAttachments().clear();
//									System.out.println("selected..");
//								email.DropAttachments();
//								}
//							email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//							MailMessage message = MailMessage
//									.load(System.getProperty("user.home") + File.separator + "0.eml");
//							String path5 = f.getAbsolutePath() + File.separator + i
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//							//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//							ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//							message.save(emlStream, SaveOptions.getDefaultMhtml());
//							LoadOptions lo = new LoadOptions();
//							lo.setLoadFormat(LoadFormat.MHTML);
//							Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//							//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//							doc.save(path5+".doc", SaveFormat.DOC);
//							
//							if(email.get_NumAttachments()>0) {
//								//	email.attachment		
//									for(int b=0;b<email.get_NumAttachments();b++) {
//										
//											File f1 = new File(destination + File.separator + foldername + File.separator
//													+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//											f.mkdirs();
//
//											//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////											byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////											String str = new String(bytes, StandardCharsets.US_ASCII);
////											System.out.println(str);
//											email.SaveAttachedFile(b,f1.getAbsolutePath());
////											attachment.save(f.getAbsolutePath() + File.separator
////													+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//										
//										
//									}
//								}
//							
//							//email.
////							If email.NumAttachedMessages > 0 then
////					        '''how do I get the original email and download the file attachments if any.
////					    else
////					        Dim i as integer
////					        For i = 0 to email.NumAttachments - 1
////					            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////					        next i
////							PdfContentEditor editor = new PdfContentEditor();
////							editor.bindPdf(path5);
////							if(email.get_NumAttachments()>0) {
////								for(int b=0;b<email.get_NumAttachments();b++) {
////									
////									File f1;
////									
////									if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////										 f1=new File(System.getProperty("java.io.tmpdir"));
////										email.SaveAttachedFile(b,f1.getAbsolutePath());
////									}
////									else {
////										f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////												+ "Application Support" );
////										email.SaveAttachedFile(b,f1.getAbsolutePath());
////									}
////			
////		com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////			// Set up a new file to be added as attachment
////			FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////			// Add an attachment to document's attachment collection
////			pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////			// Save the updated document
////			pdfDocument.save(path5);
////			File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////			f1.delete();
////			f2.delete();
////								}
////							}
//							
////							if (message.getAttachments().size() > 0) {
////								
//	//
//////								MapiMessage msg = MapiMessage.fromMailMessage(message);
//////								PdfContentEditor editor = new PdfContentEditor();
//////								editor.bindPdf(path5);
//	////
//////								for (MapiAttachment attachment : msg.getAttachments()) {
//////									if (!(attachment.getExtension() == null)) {
//////									ByteArrayOutputStream eml = new ByteArrayOutputStream();
//	////
//////									attachment.save(eml);
//////									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////											attachment.getDisplayName(), "");
//////									}
//////								}
////							}
//							//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////							lblNewLabel_3.setText("count " + i);
//
////							lblNewLabel_3_2.setText("Total Count :" + countdest++);
//							Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//									+ fname + "   Extarcting messsage " + message.getSubject());
//							Main_Frame.count_destination++;
//						} catch (Exception e) {
//							e.printStackTrace();
//							continue;
//						}
//						i++;
//					}
//
//				}
//			}
//				j++;
//			}
//			
//		}
//		else if(filetype.equalsIgnoreCase("docx")) {
//
//			while (j < mboxes.get_Count()) {
//				if (Main_Frame.stop) {
//					break;
//				}
//				String a=mboxes.getName(j).replaceAll("/", "\\\\");
//				a = a.replaceAll("[\\[\\]\\(\\)]", "");
//				System.out.println("printing a---"+a);
//				
//				if(pstfolderlist.contains(a) ) {
//				
//				
//				System.out.println("41-----"+mboxes.getName(j));
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+foldername);
//			
//
////				lblNewLabel_3_1.setText("Folder Name :" + foldername);
//				CkEmailBundle bundle;
//				CkMessageSet messageSet;
//				// We can choose to fetch UIDs or sequence numbers.
//				boolean fetchUids = true;
//				success = imap.SelectMailbox(mboxes.getName(j));
//				if (success != true) {
//					System.out.println(imap.lastErrorText());
//					// return;
//				}
//				// Get the message IDs of all the emails in the mailbox
//				messageSet = imap.Search("ALL", fetchUids);
//				if (imap.get_LastMethodSuccess() == true) {
//					if (imap.get_LastMethodSuccess() == false) {
//						System.out.println(imap.lastErrorText());
//						// return;
//					}
//
//					//	bundle = imap.FetchBundle(messageSet);
//					if (imap.get_LastMethodSuccess() == false) {
//						System.out.println(imap.lastErrorText());
//						// return;
//					}
//
//					int numEmails = imap.get_NumMessages();
//					if(numEmails>0) {
//						 f = new File(destination+ File.separator+ foldername);
//							f.mkdirs();
//						}
//					 boolean bUid = false;
//					int i = 1;
//					System.out.println(numEmails);
//					while (i <= numEmails) {
//						try {
//							if (Main_Frame.stop) {
//								break;
//							}
//							System.out.println(i);
//							CkEmail email = imap.FetchSingle(i,bUid);
//							if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//								//	message.getAttachments().clear();
//									System.out.println("selected..");
//								email.DropAttachments();
//								}
//							email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//							MailMessage message = MailMessage
//									.load(System.getProperty("user.home") + File.separator + "0.eml");
//							String path5 = f.getAbsolutePath() + File.separator + i
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//							//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//							ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//							message.save(emlStream, SaveOptions.getDefaultMhtml());
//							LoadOptions lo = new LoadOptions();
//							lo.setLoadFormat(LoadFormat.MHTML);
//							Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//							//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//							doc.save(path5+".docx", SaveFormat.DOCX);
//							
//							if(email.get_NumAttachments()>0) {
//								//	email.attachment		
//									for(int b=0;b<email.get_NumAttachments();b++) {
//										
//											File f1 = new File(destination + File.separator + foldername + File.separator
//													+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//											f.mkdirs();
//
//											//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////											byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////											String str = new String(bytes, StandardCharsets.US_ASCII);
////											System.out.println(str);
//											email.SaveAttachedFile(b,f1.getAbsolutePath());
////											attachment.save(f.getAbsolutePath() + File.separator
////													+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//										
//										
//									}
//								}
//							//email.
////							If email.NumAttachedMessages > 0 then
////					        '''how do I get the original email and download the file attachments if any.
////					    else
////					        Dim i as integer
////					        For i = 0 to email.NumAttachments - 1
////					            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////					        next i
////							PdfContentEditor editor = new PdfContentEditor();
////							editor.bindPdf(path5);
////							if(email.get_NumAttachments()>0) {
////								for(int b=0;b<email.get_NumAttachments();b++) {
////									
////									File f1;
////									
////									if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////										 f1=new File(System.getProperty("java.io.tmpdir"));
////										email.SaveAttachedFile(b,f1.getAbsolutePath());
////									}
////									else {
////										f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////												+ "Application Support" );
////										email.SaveAttachedFile(b,f1.getAbsolutePath());
////									}
////			
////		com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////			// Set up a new file to be added as attachment
////			FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////			// Add an attachment to document's attachment collection
////			pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////			// Save the updated document
////			pdfDocument.save(path5);
////			File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////			f1.delete();
////			f2.delete();
////								}
////							}
//							
////							if (message.getAttachments().size() > 0) {
////								
//	//
//////								MapiMessage msg = MapiMessage.fromMailMessage(message);
//////								PdfContentEditor editor = new PdfContentEditor();
//////								editor.bindPdf(path5);
//	////
//////								for (MapiAttachment attachment : msg.getAttachments()) {
//////									if (!(attachment.getExtension() == null)) {
//////									ByteArrayOutputStream eml = new ByteArrayOutputStream();
//	////
//////									attachment.save(eml);
//////									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////											attachment.getDisplayName(), "");
//////									}
//////								}
////							}
//							//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////							lblNewLabel_3.setText("count " + i);
//
////							lblNewLabel_3_2.setText("Total Count :" + countdest++);
//							Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//									+ fname + "   Extarcting messsage " + message.getSubject());
//							Main_Frame.count_destination++;
//						} catch (Exception e) {
//							e.printStackTrace();
//							continue;
//						}
//						i++;
//					}
//
//				}
//			}
//				j++;
//			}
//		
//		}
//else if(filetype.equalsIgnoreCase("docm")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".docm", SaveFormat.DOCM);
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
//
//		}
//else if(filetype.equalsIgnoreCase("tiff")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".tiff", SaveFormat.TIFF);
//					
//					
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
// }
//else if(filetype.equalsIgnoreCase("png")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".png", SaveFormat.PNG);
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
// }
//else if(filetype.equalsIgnoreCase("txt")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".txt", SaveFormat.TEXT);
//					
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
//
// }
//else if(filetype.equalsIgnoreCase("bmp")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".bmp", SaveFormat.BMP);
//					
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
//
// }
//else if(filetype.equalsIgnoreCase("gif")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
////					path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".gif", SaveFormat.GIF);
//					
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
//
// }
//else if(filetype.equalsIgnoreCase("jpg")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//				//	path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".jpg", SaveFormat.JPEG);
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
// }
//else if(filetype.equalsIgnoreCase("png")) {
//
//	while (j < mboxes.get_Count()) {
//		if (Main_Frame.stop) {
//			break;
//		}
//		String a=mboxes.getName(j).replaceAll("/", "\\\\");
//		a = a.replaceAll("[\\[\\]\\(\\)]", "");
//		System.out.println("printing a---"+a);
//		
//		if(pstfolderlist.contains(a) ) {
//		
//		
//		System.out.println("41-----"+mboxes.getName(j));
//
//		String foldername = mboxes.getName(j).replace("/", File.separator);
//		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//	//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//System.out.println("51--"+destination);
//System.out.println("52--"+foldername);
//	
//
////		lblNewLabel_3_1.setText("Folder Name :" + foldername);
//		CkEmailBundle bundle;
//		CkMessageSet messageSet;
//		// We can choose to fetch UIDs or sequence numbers.
//		boolean fetchUids = true;
//		success = imap.SelectMailbox(mboxes.getName(j));
//		if (success != true) {
//			System.out.println(imap.lastErrorText());
//			// return;
//		}
//		// Get the message IDs of all the emails in the mailbox
//		messageSet = imap.Search("ALL", fetchUids);
//		if (imap.get_LastMethodSuccess() == true) {
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			//	bundle = imap.FetchBundle(messageSet);
//			if (imap.get_LastMethodSuccess() == false) {
//				System.out.println(imap.lastErrorText());
//				// return;
//			}
//
//			int numEmails = imap.get_NumMessages();
//			if(numEmails>0) {
//				 f = new File(destination+ File.separator+ foldername);
//					f.mkdirs();
//				}
//			 boolean bUid = false;
//			int i = 1;
//			System.out.println(numEmails);
//			while (i <= numEmails) {
//				try {
//					if (Main_Frame.stop) {
//						break;
//					}
//					System.out.println(i);
//					CkEmail email = imap.FetchSingle(i,bUid);
//					if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
//						//	message.getAttachments().clear();
//							System.out.println("selected..");
//						email.DropAttachments();
//						}
//					email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
//					MailMessage message = MailMessage
//							.load(System.getProperty("user.home") + File.separator + "0.eml");
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ Main_Frame.getRidOfIllegalFileNameCharacters(Main_Frame.namingconventionmail(message));
//					//message.save(path5 + ".eml", SaveOptions.getDefaultEml());
//					ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//					message.save(emlStream, SaveOptions.getDefaultMhtml());
//					LoadOptions lo = new LoadOptions();
//					lo.setLoadFormat(LoadFormat.MHTML);
//					Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//				//	path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//					doc.save(path5+".png", SaveFormat.PNG);
//					
//					if(email.get_NumAttachments()>0) {
//						//	email.attachment		
//							for(int b=0;b<email.get_NumAttachments();b++) {
//								
//									File f1 = new File(destination + File.separator + foldername + File.separator
//											+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
//									f.mkdirs();
//
//									//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//
////									byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
////									String str = new String(bytes, StandardCharsets.US_ASCII);
////									System.out.println(str);
//									email.SaveAttachedFile(b,f1.getAbsolutePath());
////									attachment.save(f.getAbsolutePath() + File.separator
////											+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//								
//							}
//						}
//					//email.
////					If email.NumAttachedMessages > 0 then
////			        '''how do I get the original email and download the file attachments if any.
////			    else
////			        Dim i as integer
////			        For i = 0 to email.NumAttachments - 1
////			            email.SaveAttachedFile(j, \"C:\FileAttachments\\")
////			        next i
////					PdfContentEditor editor = new PdfContentEditor();
////					editor.bindPdf(path5);
////					if(email.get_NumAttachments()>0) {
////						for(int b=0;b<email.get_NumAttachments();b++) {
////							
////							File f1;
////							
////							if (System.getProperty("os.name").toLowerCase().contains("windows")) {
////								 f1=new File(System.getProperty("java.io.tmpdir"));
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////							else {
////								f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
////										+ "Application Support" );
////								email.SaveAttachedFile(b,f1.getAbsolutePath());
////							}
////	
////com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path5);
////
////	// Set up a new file to be added as attachment
////	FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");
////
////	// Add an attachment to document's attachment collection
////	pdfDocument.getEmbeddedFiles().add(fileSpecification);
////
////	// Save the updated document
////	pdfDocument.save(path5);
////	File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
////	f1.delete();
////	f2.delete();
////						}
////					}
//					
////					if (message.getAttachments().size() > 0) {
////						
////
//////						MapiMessage msg = MapiMessage.fromMailMessage(message);
//////						PdfContentEditor editor = new PdfContentEditor();
//////						editor.bindPdf(path5);
//////
//////						for (MapiAttachment attachment : msg.getAttachments()) {
//////							if (!(attachment.getExtension() == null)) {
//////							ByteArrayOutputStream eml = new ByteArrayOutputStream();
//////
//////							attachment.save(eml);
//////							editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
//////									attachment.getDisplayName(), "");
//////							}
//////						}
////					}
//					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
////					lblNewLabel_3.setText("count " + i);
//
////					lblNewLabel_3_2.setText("Total Count :" + countdest++);
//					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
//							+ fname + "   Extarcting messsage " + message.getSubject());
//					Main_Frame.count_destination++;
//				} catch (Exception e) {
//					e.printStackTrace();
//					continue;
//				}
//				i++;
//			}
//
//		}
//	}
//		j++;
//	}
//
// }
		/*--------------------------------------------------------------------------------------------------------------------*/
//		String naming_convention = Main_Frame.getRidOfIllegalFileNameCharacters(mf.namingconventionmail(message));
//		String path5 = destination_path + File.separator + path + File.separator + naming_convention + "_"
//				+ Main_Frame.count_destination;
//		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
//
//		message.save(emlStream, SaveOptions.getDefaultMhtml());
//		LoadOptions lo = new LoadOptions();
//		lo.setLoadFormat(LoadFormat.MHTML);
//		try {
//			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
//			if (filetype.equalsIgnoreCase("PDF")) {
//				path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
//				try {
//					doc.save(path5, SaveFormat.PDF);
//				} catch (Exception e1) {
//					// TODO Auto-generated catch block
//					e1.printStackTrace();
//				}
				
	//}

//	static String namingconventionmail(MailMessage msg) {
//		String filename = null;
//		String frm;
//		try {
//			frm = getRidOfIllegalFileNameCharacters(msg.getFrom().toString());
//		} catch (Exception ep) {
//			frm = "";
//		}
//		if (frm != null) {
//
//		} else {
//			frm = "";
//		}
//
//		if (frm.length() > 20) {
//			frm = frm.substring(0, 20);
//		}
//
//		String sub;
//		try {
//			sub = getRidOfIllegalFileNameCharacters(msg.getSubject());
//		} catch (Exception ep) {
//			sub = "";
//		}
//
//		if (sub != null) {
//
//		} else {
//			sub = "";
//		}
//
//		if (sub.length() > 40) {
//			sub = sub.substring(0, 40);
//		}
//
//		String dstr = "";
//		Date d;
//
//		try {
//			d = msg.getDate();
//			Calendar cal = Calendar.getInstance();
//			cal.setTime(d);
//			DecimalFormat formatter = new DecimalFormat("00");
//
//			int date = cal.get(Calendar.DATE);
//			String dateformate = formatter.format(date);
//
//			int month = cal.get(Calendar.MONTH);
//			month++;
//			String monthformate = formatter.format(month);
//
//			int year = cal.get(Calendar.YEAR);
//
//			dstr = dateformate + "-" + monthformate + "-" + year;
//
//		} catch (Exception ep) {
//			dstr = "";
//		}
//
//		filename = sub;
//
//		filename = getRidOfIllegalFileNameCharacters(filename);
//		return filename;
//	}

//	static String getRidOfIllegalFileNameCharacters(String strName) {
//		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
//				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
//				.replace("//s", "").replace("\"", "");
//		if (strLegalName.length() >= 80) {
//			strLegalName = strLegalName.substring(0, 80);
//		}
//		return strLegalName;
//	}
}
public boolean MailConvert(String path,MailMessage message,String filetype,CkEmail email,File f)
{
	boolean return1=false;
	if(filetype.equalsIgnoreCase("pdf")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			path = path.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path, SaveFormat.PDF);
	
			if(email.get_NumAttachments()>0) {
			//	email.attachment		
				for(int b=0;b<email.get_NumAttachments();b++) {
					if (Main_Frame.chckbxSavePdfAttachment.isSelected()) {
						File f1 = new File(f.getAbsolutePath() + File.separator
								+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
						f1.mkdirs();


						email.SaveAttachedFile(b,f1.getAbsolutePath());

					}
					else {
					File f1;
				//	email.date
					if (System.getProperty("os.name").toLowerCase().contains("windows")) {
						 f1=new File(System.getProperty("java.io.tmpdir"));
						email.SaveAttachedFile(b,f1.getAbsolutePath());
					}
					else {
						f1= new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
								+ "Application Support" );
						email.SaveAttachedFile(b,f1.getAbsolutePath());
					}

com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(path);

//Set up a new file to be added as attachment
FileSpecification fileSpecification = new FileSpecification(f1.getAbsolutePath()+File.separator+email.getAttachmentFilename(b), "");

//Add an attachment to document's attachment collection
pdfDocument.getEmbeddedFiles().add(fileSpecification);

//Save the updated document
pdfDocument.save(path);
File f2=new File(f1.getAbsoluteFile()+File.separator+email.getAttachmentFilename(b));
f1.delete();
f2.delete();
pdfDocument.close();
				}
				}
			}
	return1= true;
		} catch (Exception e) {
			
			e.printStackTrace();
			return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("doc")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			//message.setDate(null);
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".doc", SaveFormat.DOC);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
			} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
			}
	}
	if(filetype.equalsIgnoreCase("docx")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".docx", SaveFormat.DOCX);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
			} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
			}
	}
	if(filetype.equalsIgnoreCase("docm")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".docm", SaveFormat.DOCM);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("gif")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".gif", SaveFormat.GIF);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("jpg")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".jpg", SaveFormat.JPEG);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("txt")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".txt", SaveFormat.TEXT);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
		return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
	return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("tiff")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".tiff", SaveFormat.TIFF);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
			} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
			}
	}
	if(filetype.equalsIgnoreCase("png")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".png", SaveFormat.PNG);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
		}
	}
	if(filetype.equalsIgnoreCase("bmp")) {
		try {
			ByteArrayOutputStream emlStream = new ByteArrayOutputStream();
			message.save(emlStream, SaveOptions.getDefaultMhtml());
			LoadOptions lo = new LoadOptions();
			lo.setLoadFormat(LoadFormat.MHTML);
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			//path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
			doc.save(path+".bmp", SaveFormat.BMP);
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath()+ File.separator
									+ "Attachment" + File.separator + message.getSubject()+"_"	+ Main_Frame.count_destination);
							f1.mkdirs();

							//String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");

//							byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//							String str = new String(bytes, StandardCharsets.US_ASCII);
//							System.out.println(str);
							email.SaveAttachedFile(b,f1.getAbsolutePath());
//							attachment.save(f.getAbsolutePath() + File.separator
//									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

						
						
					}
				}
			return1= true;
		} catch (Exception e) {
		
			e.printStackTrace();
			return1= false;
		}
	}
	return return1;
}


}
