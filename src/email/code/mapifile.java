package email.code;

import java.io.File;
import java.util.Date;

import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiMessage;
import com.aspose.email.SaveOptions;

public class mapifile implements Runnable {
	private Main_Frame mf;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	private String path = "";
	long count_destination;

	private MapiMessage message = null;
	private Date reciveddate = null;
	private Date startdate = null;
	private Date enddate = null;
	private Boolean Filter = null;
	long k;

	public mapifile(Main_Frame mf, String filetype, String destination_path, String path, long count_destination,
			MapiMessage message, Date reciveddate, boolean Filter, Date startdate, Date enddate, String filepath) {
		this.mf = mf;
		this.filetype = filetype;
		this.destination_path = destination_path;
		this.path = path;
		this.count_destination = count_destination;
		this.message = message;
		this.reciveddate = reciveddate;
		mf.stop = false;
		this.Filter = Filter;
		this.startdate = startdate;
		this.filepath = filepath;
		this.enddate = enddate;

	}

	public void run() {
		mapidupfile(message, reciveddate);
	}

	void mapidupfile(MapiMessage message, Date reciveddate) {

		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!mf.listduplicacy.contains(input)) {

				mf.listduplicacy.add(input);

				if (Filter) {
					if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
						Mapimessage_file(message);
						Main_Frame.foldermessagecount++;
					} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
						Mapimessage_file(message);
						Main_Frame.foldermessagecount++;
					}
				} else {

					Mapimessage_file(message);
					Main_Frame.foldermessagecount++;
				}
			} else {
				System.out.println(" duplicate message");
				System.out.println(input);

			}
		} else {
			System.out.println(Filter);
			if (Filter) {
				if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
					Mapimessage_file(message);
					Main_Frame.foldermessagecount++;
				} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
					Mapimessage_file(message);
					Main_Frame.foldermessagecount++;
				}
			} else {
				Mapimessage_file(message);
				Main_Frame.foldermessagecount++;
			}
		}

	}

	public void Mapimessage_file(MapiMessage message) {
		
		String subname = "";
		if (Main_Frame.fileoption.equals("OLM files (.olm)")) {
			subname = mf.namingconventionmapi(message, mf.c1);
		} else if (Main_Frame.fileoption.equals("EML File (.eml)") || Main_Frame.fileoption.equals("EMLX File (.emlx)")
				|| Main_Frame.fileoption.equals("Message File (.msg)") || Main_Frame.fileoption.equals("Maildir")) {
			subname = mf.namingconventionmapi(message, new File(filepath));
		} else {
			subname = mf.namingconventionmapi(message);
		}
		String finalpath = destination_path + File.separator + path + File.separator +Main_Frame.getRidOfIllegalFileNameCharacters(subname) ;
		try {
			File fu = null;
			if (filetype.equalsIgnoreCase("EML")) {
				try {

					fu = new File(finalpath + ".eml");
					if (fu.isFile()) {

						message.save(fu.getAbsolutePath() + "_" + Main_Frame.count_destination + ".eml",
								SaveOptions.getDefaultEml());
					} else {
						message.save(finalpath + ".eml", SaveOptions.getDefaultEml());

					}
				} catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ +Main_Frame.count_destination + ".eml", SaveOptions.getDefaultEml());
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("HTML")) {

				try {

					fu = new File(finalpath + ".html");
					if (fu.isFile()) {

						message.save(finalpath + "_" + Main_Frame.count_destination + ".html",
								SaveOptions.getDefaultHtml());
					} else {
						message.save(finalpath + ".html", SaveOptions.getDefaultHtml());

					}

					new File(destination_path + File.separator + path + File.separator + "Attachment").mkdirs();

					for (int j = 0; j < message.getAttachments().size(); j++) {
						MapiAttachment att = message.getAttachments().get_Item(j);

						String s = Main_Frame.getFileExtension(att.getDisplayName());
						String attFileName = Main_Frame
								.getRidOfIllegalFileNameCharacters(att.getDisplayName().replace("." + s, ""));

						att.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + attFileName);

					}

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ +Main_Frame.count_destination + ".html", SaveOptions.getDefaultHtml());
				}

			} else if (filetype.equalsIgnoreCase("MSG")) {

				try {
					fu = new File(finalpath + ".msg");
					if (fu.isFile()) {
						message.save(finalpath + "_" + Main_Frame.count_destination + ".msg",
								SaveOptions.getDefaultMsg());

					} else {

						message.save(finalpath + ".msg", SaveOptions.getDefaultMsg());
					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ +Main_Frame.count_destination + ".msg", SaveOptions.getDefaultMsg());
				}

			} else if (filetype.equalsIgnoreCase("EMLX")) {
				MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
				try {

					fu = new File(finalpath + ".emlx");
					if (fu.isFile()) {
//message.getDisplayTo();
	
						message.save(finalpath + "_" + Main_Frame.count_destination + ".emlx",
								SaveOptions.createSaveOptions(messagesavetype));
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
					SaveOptions.createSaveOptions(messagesavetype));

					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
							SaveOptions.createSaveOptions(messagesavetype));
				}
			} else if (filetype.equalsIgnoreCase("MHTML")) {

				try {
					fu = new File(finalpath + ".mhtml");
					if (fu.isFile()) {
						message.save(finalpath + "_" + Main_Frame.count_destination + ".mhtml",
								SaveOptions.getDefaultMhtml());

					} else {

						message.save(finalpath + ".mhtml", SaveOptions.getDefaultMhtml());
					}
					new File(destination_path + File.separator + path + File.separator + "Attachment").mkdirs();

					for (int j = 0; j < message.getAttachments().size(); j++) {
						MapiAttachment att = message.getAttachments().get_Item(j);

						String s = Main_Frame.getFileExtension(att.getDisplayName());
						String attFileName = Main_Frame
								.getRidOfIllegalFileNameCharacters(att.getDisplayName().replace("." + s, ""));

						att.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + attFileName);

					}

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator
							+ +Main_Frame.count_destination + ".mhtml", SaveOptions.getDefaultMhtml());
				}

			}
			fu = null;

			k = Main_Frame.count_destination++;

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
