package email.code;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;

import com.aspose.email.FolderInfo;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiMessage;
import com.aspose.email.MboxrdStorageWriter;
import com.chilkatsoft.CkEmail;
import com.chilkatsoft.CkEmailBundle;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkMailboxes;
import com.chilkatsoft.CkMessageSet;

public class ConvertToMbox {
	static CkImap imap;
	 ConvertToMbox(CkImap imap1,String destination,String filetype,ArrayList<String> pstfolderlist,Date startdate,Date enddate) {		
		File f = null;
this.imap=imap1;

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
			// FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);

//			File f = new File(file.getAbsolutePath() + foldername);
//			f.mkdirs();

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
				MboxrdStorageWriter wr=null;
			if(numEmails>0) {
			 f = new File(destination+ File.separator+ foldername);
				f.mkdirs();
				
				wr = new MboxrdStorageWriter(destination + File.separator + foldername + File.separator
						+ fname, false);
			}
		    boolean bUid = false;
				int i = 1;
				System.out.println(numEmails);
				while (i <= numEmails) {
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
						if (Main_Frame.chckbxMigrateOrBackup.isSelected()) {
							//	message.getAttachments().clear();
								System.out.println("selected..");
							email.DropAttachments();
							}
						email.SaveEml(System.getProperty("user.home") + File.separator + "0.eml");
						MailMessage message = MailMessage
								.load(System.getProperty("user.home") + File.separator + "0.eml");
				
						Date reciveddate=message.getDate();
						//=========
						if(Main_Frame.chckbxRemoveDuplicacy.isSelected()) {


							String input = Main_Frame.duplicacymail(message);

							if (!Main_Frame.listduplicacy.contains(input)) {

								Main_Frame.listduplicacy.add(input);

								if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
									if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
										wr.writeMessage(message);
										Main_Frame.count_destination++;
									} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
										wr.writeMessage(message);
										Main_Frame.count_destination++;
									}
								}else {
									wr.writeMessage(message);
									Main_Frame.count_destination++;
								}
								
								} else {
								System.out.println(" duplicate message");
								System.out.println(input);

							}
						
						}else {	
							if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
							if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
								wr.writeMessage(message);
								Main_Frame.count_destination++;
							} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
								wr.writeMessage(message);
								Main_Frame.count_destination++;
							}
						}else {
							wr.writeMessage(message);
							Main_Frame.count_destination++;
						}
						}	
						//=========
						
						
					
					
//						String path5 = f.getAbsolutePath() + File.separator + i
//								+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//						message.save(path5 + ".eml", SaveOptions.getDefaultEml());
						//folerinfo.addMessage(MapiMessage.fromMailMessage(message));
//						lblNewLabel_3.setText("count " + i);

//						lblNewLabel_3_2.setText("Total Count :" + countdest++);
						Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
								+ fname + "   Extarcting messsage " + message.getSubject());
						//Main_Frame.count_destination++;
						if(Main_Frame.chckbxDeleteEmailFrom.isSelected()) {
						//	if(mail_convert) {
								success = imap.SetMailFlag(email,"Deleted",1);
					//	}
						
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
//					if(Main_Frame.chckbxDeleteEmailFrom.isSelected()) {
//						
//						imap.SetFlags(messageSet,"Deleted",i);	
//					
//					
//				}
					i++;
				
				}

			}
		}
			j++;
		}

	}
	
}
