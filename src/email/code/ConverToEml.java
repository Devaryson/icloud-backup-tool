package email.code;

import java.io.File;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

import com.aspose.email.FolderInfo;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiMessage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.chilkatsoft.CkEmail;
import com.chilkatsoft.CkEmailBundle;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkMailboxes;
import com.chilkatsoft.CkMessageSet;

public class ConverToEml {
	static CkImap imap;
	ConverToEml(CkImap imap1,String destination,String filetype,ArrayList<String> pstfolderlist,Date startdate,Date enddate){
//imap.
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
		//if(filetype.equalsIgnoreCase("eml")) {
		while (j < mboxes.get_Count()) {
			if (Main_Frame.stop) {
				break;
			}
			String a=mboxes.getName(j).replaceAll("/", "\\\\");
		//	a = a.replaceAll("[^a-zA-Z0-9]", "");
			a = a.replaceAll("[\\[\\]\\(\\)]", "");
		//	System.out.println(str);
			System.out.println("printing a---"+a);
			
			if(pstfolderlist.contains(a) ) {
			
			
			System.out.println("41-----"+mboxes.getName(j));
			

			String foldername = mboxes.getName(j).replace("/", File.separator);
		//	int fin=foldername.lastIndexOf(File.separator);
//	System.out.println("index--"+fin);
			String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
		//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
System.out.println("51--"+destination);
System.out.println("52--"+fname);

//mboxes.i
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
						
						
						
						System.out.println(i);
						CkEmail email =imap.FetchSingle(i,bUid);
						// email.get_EmailDate(null);
					//	System.out.println("118--");
								//bundle.GetEmail(i);
						if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
							//	message.getAttachments().clear();
								System.out.println("selected..");
							email.DropAttachments();
							}
						
						email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
						MailMessage message = MailMessage
								.load(System.getProperty("user.home") + File.separator + "0.eml");
						String path5 = f.getAbsolutePath() + File.separator + i
								+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
						Date reciveddate=message.getDate();
						System.out.println("date--"+message.getDate().toString());
						//-----------------
						if(Main_Frame.chckbxRemoveDuplicacy.isSelected()) {


							String input = Main_Frame.duplicacymail(message);

							if (!Main_Frame.listduplicacy.contains(input)) {

								Main_Frame.listduplicacy.add(input);

								if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
									if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
										mail_convert=MailConvert(path5, message, filetype, email,f);
										Main_Frame.count_destination++;
									} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
										mail_convert=MailConvert(path5, message, filetype, email,f);
										Main_Frame.count_destination++;
									}
								} else {
									mail_convert=MailConvert(path5, message, filetype, email,f);
									Main_Frame.count_destination++;
								}
								} else {
								System.out.println(" duplicate message");
								System.out.println(input);

							}
						
						}else {
							if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
							if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
								mail_convert=MailConvert(path5, message, filetype, email,f);
								Main_Frame.count_destination++;
							} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
								mail_convert=MailConvert(path5, message, filetype, email,f);
								Main_Frame.count_destination++;
							}
						} else {
							mail_convert=MailConvert(path5, message, filetype, email,f);
							Main_Frame.count_destination++;
						}
							}	
						//-----------------
						
						
						
					//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
//						lblNewLabel_3.setText("count " + i);

//						lblNewLabel_3_2.setText("Total Count :" + countdest++);
						
						Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
								+ fname + "   Extarcting messsage " + message.getSubject());
						//Main_Frame.count_destination++;
						if(Main_Frame.chckbxDeleteEmailFrom.isSelected()) {
							if(mail_convert) {
								success = imap.SetMailFlag(email,"Deleted",1);
						}
						
					}
					} catch (Exception e) {
					//	System.out.println("117-----"+numEmails);
					//	e.printStackTrace();
					//	continue;
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
				
					i++;
				}

			}
		}
			j++;
		}
	//}
//		else if(filetype.equalsIgnoreCase("emlx")) {
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
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//			//	int fin=foldername.lastIndexOf(File.separator);
////		System.out.println("index--"+fin);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+fname);
//
//	//mboxes.i
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
//									+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//							MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
//							message.save(path5 + ".emlx", SaveOptions.createSaveOptions(messagesavetype));
//						//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
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
//		else if(filetype.equalsIgnoreCase("msg")) {
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
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//			//	int fin=foldername.lastIndexOf(File.separator);
////		System.out.println("index--"+fin);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+fname);
//
//	//mboxes.i
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
//									+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//							message.save(path5 + ".msg", SaveOptions.getDefaultMsg());
//						//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
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
//		else if(filetype.equalsIgnoreCase("html")) {
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
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//			//	int fin=foldername.lastIndexOf(File.separator);
////		System.out.println("index--"+fin);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+fname);
//
//	//mboxes.i
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
//									+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//							message.save(path5 + ".html", SaveOptions.getDefaultHtml());
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
//						//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
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
//		else if(filetype.equalsIgnoreCase("mhtml")) {
//
//			while (j < mboxes.get_Count()) {
//				if (Main_Frame.stop) {
//					break;
//				}
//				String a=mboxes.getName(j).replaceAll("/", "\\\\");
//				a = a.replaceAll("[\\[\\]\\(\\)]", "");
//				a = a.replaceAll("[^a-zA-Z0-9]", " ");
//				System.out.println("printing a---"+a);
//				
//				if(pstfolderlist.contains(a) ) {
//				
//				
//				System.out.println("41-----"+mboxes.getName(j));
//				
//
//				String foldername = mboxes.getName(j).replace("/", File.separator);
//			//	int fin=foldername.lastIndexOf(File.separator);
////		System.out.println("index--"+fin);
//				String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
//			//	 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);
//	System.out.println("51--"+destination);
//	System.out.println("52--"+fname);
//
//	//mboxes.i
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
//									+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//							message.save(path5 + ".mhtml", SaveOptions.getDefaultMhtml());
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
//						//	folerinfo.addMessage(MapiMessage.fromMailMessage(message));
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
	}
	
	public boolean MailConvert(String path5,MailMessage message,String filetype,CkEmail email,File f) {
	//	boolean return1=false;
		if(filetype.equalsIgnoreCase("eml")) {
		message.save(path5 + ".eml", SaveOptions.getDefaultEml());
	}
		else if(filetype.equalsIgnoreCase("emlx")) {
			MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
			message.save(path5 + ".emlx", SaveOptions.createSaveOptions(messagesavetype));
		}
		else if(filetype.equalsIgnoreCase("msg")) {
			message.save(path5 + ".msg", SaveOptions.getDefaultMsg());
		}
		else if(filetype.equalsIgnoreCase("html")) {
			message.save(path5 + ".html", SaveOptions.getDefaultHtml());
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath() + File.separator
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
		}
		else if(filetype.equalsIgnoreCase("mhtml")) {
			message.save(path5 + ".html", SaveOptions.getDefaultHtml());
			
			if(email.get_NumAttachments()>0) {
				//	email.attachment		
					for(int b=0;b<email.get_NumAttachments();b++) {
						
							File f1 = new File(f.getAbsolutePath() + File.separator
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
		}
	return true;	
	}
		
	
	static String namingconventionmail(MailMessage msg) {
		String filename = null;
		String frm;
		try {
			frm = getRidOfIllegalFileNameCharacters(msg.getFrom().toString());
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = getRidOfIllegalFileNameCharacters(msg.getSubject());
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}

		String dstr = "";
		Date d;

		try {
			d = msg.getDate();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);

			dstr = dateformate + "-" + monthformate + "-" + year;

		} catch (Exception ep) {
			dstr = "";
		}

		filename = sub;

		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}
}
