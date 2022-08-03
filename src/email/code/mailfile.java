package email.code;

import java.io.File;
import java.util.Date;

import com.aspose.email.Attachment;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.SaveOptions;

public class mailfile implements Runnable {
	private Main_Frame mf;
	private String filetype = "";
	private String destination_path = "";
	private String path = "";
	private String filepath = "";

	private MailMessage message = null;
	private Date reciveddate = null;
	private Date startdate;
	private boolean Filter;
	private Date enddate;
	long k;
	public mailfile(Main_Frame mf, String filetype, String destination_path, String path, MailMessage message,
			Date reciveddate, boolean Filter, Date startdate, Date enddate,String filepath) {
		this.mf = mf;
		this.filetype = filetype;
		this.destination_path = destination_path;
		this.path = path;
		this.filepath = filepath;
		this.message = message;
		this.reciveddate = reciveddate;
		this.Filter = Filter;
		this.startdate = startdate;
		this.enddate = enddate;

	}

	public void run() {
		mailfile1(message, reciveddate);
	}

	void mailfile1(MailMessage message, Date reciveddate) {

		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!mf.listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				mf.listduplicacy.add(input);

				if (Filter) {
					if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
						Mailmessage_file(message);

					} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
						Mailmessage_file(message);

					}
				} else {
					Mailmessage_file(message);

				}
			} else {
				System.out.println(" duplicate message");
				System.out.println(input);

			}
		} else {
			if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
					Mailmessage_file(message);

				} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
					Mailmessage_file(message);

				}
			} else {
				Mailmessage_file(message);

			}
		}

	}

	public void Mailmessage_file(MailMessage message) {
		try {
			String subname = "";
			if (Main_Frame.fileoption.equals("EML File (.eml)") || Main_Frame.fileoption.equals("EMLX File (.emlx)")
					|| Main_Frame.fileoption.equals("Message File (.msg)") || Main_Frame.fileoption.equals("Maildir")) {
				subname = mf.namingconventionmail(message,new File(filepath));
			} else {
				subname = mf.namingconventionmail(message);
			}
		
			String finalpath = destination_path + File.separator + path + File.separator + subname;
			File fu = null;
			if (filetype.equalsIgnoreCase("EML")) {
				try {

					fu = new File(finalpath + ".eml");
					if (fu.isFile()) {

						message.save(finalpath+"_"+
								+ Main_Frame.count_destination + ".eml", SaveOptions.getDefaultEml());
					} else {
						message.save(finalpath + ".eml",
								SaveOptions.getDefaultEml());

					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ Main_Frame.count_destination + ".eml", SaveOptions.getDefaultEml());
				}
			} else if (filetype.equalsIgnoreCase("HTML")) {

				try {

					fu = new File(finalpath+ ".html");
					if (fu.isFile()) {

						message.save(finalpath+"_"
								+ Main_Frame.count_destination + ".html", SaveOptions.getDefaultHtml());
					} else {
						message.save(finalpath+ ".html",
								SaveOptions.getDefaultHtml());

					}
					new File(destination_path + File.separator + path + File.separator + "Attachment").mkdirs();

					for (int j = 0; j < message.getAttachments().size(); j++) {
						Attachment att = (Attachment) message.getAttachments().get_Item(j);

						String s = Main_Frame.getFileExtension(att.getName());
						String attFileName = Main_Frame
								.getRidOfIllegalFileNameCharacters(att.getName().replace("." + s, ""));

						att.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + attFileName);

					}

				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ Main_Frame.count_destination + ".html", SaveOptions.getDefaultHtml());
				}

			} else if (filetype.equalsIgnoreCase("MSG")) {

				try {
					fu = new File(finalpath + ".msg");
					if (fu.isFile()) {
						message.save(finalpath+"_"+
								+ Main_Frame.count_destination + ".msg", SaveOptions.getDefaultMsg());

					} else {

						message.save(finalpath+ ".msg",
								SaveOptions.getDefaultMsg());
					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ Main_Frame.count_destination + ".msg", SaveOptions.getDefaultMsg());
				}

			} else if (filetype.equalsIgnoreCase("EMLX")) {
				MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
				try {

					fu = new File(finalpath+ ".emlx");
					if (fu.isFile()) {

						message.save(
								finalpath+"_"+
										+ Main_Frame.count_destination + ".emlx",
								SaveOptions.createSaveOptions(messagesavetype));
					} else {
						message.save(finalpath+ ".emlx",
								SaveOptions.createSaveOptions(messagesavetype));

					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
							SaveOptions.createSaveOptions(messagesavetype));
				}
			} else if (filetype.equalsIgnoreCase("MHTML")) {

				try {
					fu = new File(finalpath+ ".mhtml");
					if (fu.isFile()) {
						message.save(finalpath+"_"+
								+ Main_Frame.count_destination + ".mhtml", SaveOptions.getDefaultMhtml());

					} else {

						message.save(finalpath+"_"+ ".mhtml",
								SaveOptions.getDefaultMhtml());
					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ Main_Frame.count_destination + ".mhtml", SaveOptions.getDefaultMhtml());
				}

			}
			fu = null;
		k=	Main_Frame.count_destination++;

		} catch (Exception e) {
			return;
		}

	}
}
