package email.code;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiMessage;
import com.aspose.email.PersonalStorage;
import com.chilkatsoft.CkEmail;
import com.chilkatsoft.CkEmailBundle;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkMailboxes;
import com.chilkatsoft.CkMessageSet;

public  class ConvertToPst {
	//static  int count=0;
	static CkImap imap;
 ConvertToPst(CkImap imap1,PersonalStorage pst,ArrayList<String> pstfolderlist,Date startdate,Date enddate,File fileck){
	this.imap=imap1;
	 boolean success;
	 int pstindex=0;
	CkMailboxes mboxes = imap.ListMailboxes("", "*");

	if (imap.get_LastMethodSuccess() == false) {
		System.out.println(imap.lastErrorText());
		return;
	}
	
//	textField.setEnabled(false);
//	textField_1.setEnabled(false);
//	lblNewLabel_2.setText("Connected");
	long countdest = 0;
//	File file = new File(
//			System.getProperty("user.home") + "/Desktop" + File.separator + "Test");
	int j = 0;
	int l=0;
//	PersonalStorage pst = PersonalStorage.create(file.getAbsolutePath() + ".pst",
//			FileFormatVersion.Unicode);
	while (j < mboxes.get_Count()) {
		if (Main_Frame.stop) {
			break;
		}
		String a=mboxes.getName(j).replaceAll("/", "\\\\");
		a = a.replaceAll("[\\[\\]\\(\\)]", "");
		System.out.println("printing a---"+a);
		System.out.println("----->"+pstfolderlist);
		if(pstfolderlist.contains(a) ) {   
		
		
		System.out.println("41-----"+mboxes.getName(j));

		String foldername = mboxes.getName(j).replace("/", File.separator);
		String fname=foldername.substring(foldername.lastIndexOf(File.separator)+1);
		 FolderInfo folerinfo = pst.getRootFolder().addSubFolder(foldername, true);

//		File f = new File(file.getAbsolutePath() + foldername);
//		f.mkdirs();

//		lblNewLabel_3_1.setText("Folder Name :" + foldername);
		CkEmailBundle bundle;
		CkMessageSet messageSet;
		// We can choose to fetch UIDs or sequence numbers.
		boolean fetchUids = true;
	//	pst.get
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
//		if(numEmails>0) {
//		 f = new File(destination+ File.separator+ fname);
//			f.mkdirs();
//		}
	    boolean bUid = false;
			int i = 1;
			System.out.println(numEmails);
			
			File filechk=new File(fileck.getAbsolutePath()+".pst");
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
//					String path5 = f.getAbsolutePath() + File.separator + i
//							+ getRidOfIllegalFileNameCharacters(namingconventionmail(message));
//					message.save(path5 + ".eml", SaveOptions.getDefaultEml());
					
					long filesize=filechk.length();
					
					if(Main_Frame.chckbx_splitpst.isSelected()) {
						if(filesize>Main_Frame.maxsize) {
							pstindex++;
							pst = PersonalStorage.create(fileck.getAbsolutePath() + pstindex + ".pst",
									FileFormatVersion.Unicode);
//							pst.getRootFolder().addSubFolder(fname);
//							pst.getStore().changeDisplayName(username_p2);
							filechk = new File(fileck.getAbsolutePath() + pstindex + ".pst");

							folerinfo = pst.getRootFolder().addSubFolder(foldername, true);

						}	
					}
				
					
					
					Date reciveddate=message.getDate();
				//	System.out.println("date--"+reciveddate);
					//SimpleDateFormat dft=new SimpleDateFormat("DD MM YYYY HH:mm ");
				//  dft.format(message.getDate());
				//	message.getDate().setSeconds(0);
				//  Date date1=new SimpleDateFormat("dd MM yyyy HH:mm").parse(dft.format(message.getDate()));
					reciveddate.setSeconds(0);
					message.setDate(reciveddate);
					
					//-------
					if(Main_Frame.chckbxRemoveDuplicacy.isSelected()) {


						String input = Main_Frame.duplicacymail(message);

						if (!Main_Frame.listduplicacy.contains(input)) {

							Main_Frame.listduplicacy.add(input);

							if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
								if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
									folerinfo.addMessage(MapiMessage.fromMailMessage(message));
									Main_Frame.count_destination++;
								} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
									folerinfo.addMessage(MapiMessage.fromMailMessage(message));
									Main_Frame.count_destination++;
								}
							}else {
								folerinfo.addMessage(MapiMessage.fromMailMessage(message));
								Main_Frame.count_destination++;
							}
							
							} else {
							System.out.println(" duplicate message");
							System.out.println(input);

						}
					
					}else {	
						if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
						if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
							folerinfo.addMessage(MapiMessage.fromMailMessage(message));
							Main_Frame.count_destination++;
						} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
							folerinfo.addMessage(MapiMessage.fromMailMessage(message));
							Main_Frame.count_destination++;
						}
					}else {
					//	System.out.println("message date : "+message.getDate());
//						System.out.println("message date2 : "+message.getLocalDate());
						folerinfo.addMessage(MapiMessage.fromMailMessage(message));
						Main_Frame.count_destination++;
					}
						}	
					//-------
					
					
					
					
				
					
					
//					folerinfo.addMessage(MapiMessage.fromMailMessage(message));
//					lblNewLabel_3.setText("count " + i);

//					lblNewLabel_3_2.setText("Total Count :" + countdest++);
					Main_Frame.lbl_progressreport.setText("<html><b>   Total Message Saved Count   " +Main_Frame.count_destination + "  "
							+ fname + "   Extarcting messsage " + message.getSubject());
					if(Main_Frame.chckbxDeleteEmailFrom.isSelected()) {
//						if(mail_convert) {
							success = imap.SetMailFlag(email,"Deleted",1);
//					}
					
				}
	//				Main_Frame.count_destination++;
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
				
				i++;
			
			}

		}
	}
		j++;
	}
}
}
