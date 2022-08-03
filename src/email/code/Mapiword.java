package email.code;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.Date;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import com.aspose.email.Attachment;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;

import com.aspose.email.MapiMessage;
import com.aspose.email.MapiPropertyTag;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.facades.PdfContentEditor;
import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;

public class Mapiword implements Runnable {
	private Main_Frame mf;
	private String filetype = "";
	private String destination_path = "";
	private String path = "";

	long count_destination;

	private MapiMessage message = null;
	private Date reciveddate = null;
	private Date startdate = null;
	private Date enddate = null;
	private Boolean Filter = null;
	long k;
	private String filepath = "";

	public Mapiword(Main_Frame mf, String filetype, String destination_path, String path, long count_destination,
			String temp, MapiMessage message, Date reciveddate, Boolean Filter, Date startdate, Date enddate,
			String filepath) {
		this.mf = mf;
		this.filetype = filetype;

		this.destination_path = destination_path;
		this.path = path;
		this.filepath = filepath;
		this.count_destination = count_destination;
		this.message = message;
		this.Filter = Filter;
		this.startdate = startdate;
		this.enddate = enddate;

		this.reciveddate = reciveddate;
		mf.stop = false;
	}

	public void run() {
		mapidupword(message, reciveddate);
	}

	void mapidupword(MapiMessage message, Date reciveddate) {

		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!mf.listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				mf.listduplicacy.add(input);

				if (Filter) {
					if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
						Mapimessage_word(message);
						Main_Frame.foldermessagecount++;
					} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
						Mapimessage_word(message);
						Main_Frame.foldermessagecount++;
					}
				} else {

					Mapimessage_word(message);
					Main_Frame.foldermessagecount++;
				}
			} else {
				System.out.println(" duplicate message");
				System.out.println(input);

			}
		} else {
			if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
					Mapimessage_word(message);
					Main_Frame.foldermessagecount++;
				} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
					Mapimessage_word(message);
					Main_Frame.foldermessagecount++;
				}
			} else {
				Mapimessage_word(message);
				Main_Frame.foldermessagecount++;
			}
		}

	}

	@SuppressWarnings("unchecked")
	public void Mapimessage_word(MapiMessage message) {
		String path5 = "";
		String naming_convention = "";
		if (Main_Frame.fileoption.equals("EML File (.eml)") || Main_Frame.fileoption.equals("EMLX File (.emlx)")
				|| Main_Frame.fileoption.equals("Message File (.msg)") || Main_Frame.fileoption.equals("Maildir")) {
			naming_convention = mf.namingconventionmapi(message, new File(filepath));
			path5 = destination_path + File.separator + path + File.separator + naming_convention + "_"
					+ +Main_Frame.count_destination;

		} else {
			naming_convention = mf.namingconventionmapi(message);
			path5 = destination_path + File.separator + path + File.separator + naming_convention
					+ Main_Frame.count_destination;
		}

		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();

		message.save(emlStream, SaveOptions.getDefaultMhtml());
		LoadOptions lo = new LoadOptions();
		lo.setLoadFormat(LoadFormat.MHTML);

		MailConversionOptions options = new MailConversionOptions();
		MailMessage msg = message.toMailMessage(options);
		try {
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			if (filetype.equalsIgnoreCase("PDF")) {
				path5 = path5.replaceAll("\\p{C}", "-") + ".pdf";
				doc.save(path5, SaveFormat.PDF);

				if (message.getAttachments().size() > 0) {
					PdfContentEditor editor = new PdfContentEditor();
					editor.bindPdf(path5);

					for (MapiAttachment attachment : message.getAttachments()) {

					//	if (!(attachment.getExtension() == null)) {

							if (mf.chckbxSavePdfAttachment.isSelected()) {
								File f = new File(destination_path + File.separator + path + File.separator
										+ "Attachment" + File.separator + naming_convention);
								f.mkdirs();

								String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
								byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
								String str = new String(bytes, StandardCharsets.US_ASCII);
								System.out.println(str);

								attachment.save(f.getAbsolutePath() + File.separator
										+ Main_Frame.getRidOfIllegalFileNameCharacters(str));

							} else if (mf.chckbx_convert_pdf_to_pdf.isSelected()) {

								System.out.println("reached");

//							File f = new File(destination_path + File.separator + path + File.separator
//									+ "Attachment" + File.separator + naming_convention);
//							f.mkdirs();

								String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
								byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
								String str = new String(bytes, StandardCharsets.US_ASCII);
								str = Main_Frame.getRidOfIllegalFileNameCharacters(str);
								System.out.println(str);

								attachment.save(emlStream);

								try {

									if (str.endsWith("txt")) {
										Charset charset = Charset.forName("GB2312");

										LoadOptions loadOptions = new LoadOptions();

										loadOptions.setLoadFormat(LoadFormat.TEXT);

										loadOptions.setEncoding(charset);
										Document doc1 = new Document(new ByteArrayInputStream(emlStream.toByteArray()),
												loadOptions);
										File f = new File(mf.temppath + File.separator + str.replace("txt", ""));
										doc1.save(f.getAbsolutePath() + "pdf", SaveFormat.PDF);
										System.out.println("hi");

										editor.addDocumentAttachment(f.getAbsolutePath() + "pdf", "");
										
										f.delete();

									} else if (str.endsWith("docx")) {
										
										Charset charset = Charset.forName("GB2312");

										LoadOptions loadOptions = new LoadOptions();

										loadOptions.setLoadFormat(LoadFormat.DOCX);

										loadOptions.setEncoding(charset);
										Document doc1 = new Document(new ByteArrayInputStream(emlStream.toByteArray()),
												loadOptions);
										File f = new File(mf.temppath + File.separator + str.replace("docx", ""));
										doc1.save(f.getAbsolutePath() + "pdf", SaveFormat.PDF);
										System.out.println("hi");

										editor.addDocumentAttachment(f.getAbsolutePath() + "pdf", "");
										
										f.delete();

									} else {
										Document doc1 = new Document(new ByteArrayInputStream(emlStream.toByteArray()),
												lo);
										ByteArrayOutputStream eml = new ByteArrayOutputStream();
										doc1.save(eml, SaveFormat.PDF);
										System.out.println("hi");

										editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
												attachment.getDisplayName(), "");
									}
//									new File(f.getAbsolutePath() + File.separator
//											+ 0).delete();
								} catch (Exception e) {
									ByteArrayOutputStream eml = new ByteArrayOutputStream();

									attachment.save(eml);

									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
											attachment.getDisplayName(), "");

								}

							} else {
								ByteArrayOutputStream eml = new ByteArrayOutputStream();
								
						//	attachment.getObjectData().getProperties().getValues();

								attachment.save(eml);
								
							

								if(attachment.getProperties().containsKey(MapiPropertyTag.PR_ATTACH_EXTENSION_W))
								{
								System.out.println("Long File Name = "+ attachment.getLongFileName() + ", FileName = " + attachment.getFileName() + ", Extension = " + attachment.getExtension());
								}
								else
								{
								System.out.println("Long File Name = "+ attachment.getLongFileName() + ", FileName = " + attachment.getFileName() + ", Extension = " + attachment.getExtension());
								}
								System.out.println("Mime Tag = "+ attachment.getMimeTag());
								
								
								if (attachment.getMimeTag().contains("message")) {

									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
											attachment.getDisplayName() + ".msg", "");
								} else {
									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
											attachment.getDisplayName() , "");
								}

							}
					//	}

					}

					editor.save(path5);
					System.out.println("Embaded PDF file is created");
				}

				System.out.println("PDF Created!");

				System.out.print("save");

			} else if (filetype.equalsIgnoreCase("Json"))

			{

				try {
					JSONObject employeeDetails = new JSONObject();
					employeeDetails.put("Subject", msg.getSubject());
					employeeDetails.put("From", msg.getFrom().toString());

					if (msg.getSender() != null) {
						employeeDetails.put("Sender", msg.getSender().toString());
					} else {
						employeeDetails.put("Sender", "");
					}

					employeeDetails.put("Date", msg.getDate().toString());
					employeeDetails.put("Bcc", msg.getBcc().toString());
					employeeDetails.put("Cc", msg.getCC().toString());
					employeeDetails.put("Body", msg.getBody().toString());

					for (Attachment attachment : msg.getAttachments()) {
						new File(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + Main_Frame.getRidOfIllegalFileNameCharacters(naming_convention))
										.mkdirs();
						employeeDetails.put("Attachment", attachment.getName().toString());
						attachment.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + Main_Frame.getRidOfIllegalFileNameCharacters(naming_convention)
								+ File.separator + attachment.getName().toString());
					}

					JSONObject employeeObject = new JSONObject();
					employeeObject.put(message.getSubject(), employeeDetails);

					JSONArray employeeList = new JSONArray();
					employeeList.add(employeeObject);

					try (FileWriter file = new FileWriter(path5 + Main_Frame.count_destination + ".Json")) {

						file.write(employeeList.toJSONString());
						file.flush();

					} catch (IOException e) {
						e.printStackTrace();
					}

				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + msg.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning(
							"Exception" + e.getMessage() + "Message" + " " + msg.getDate() + System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("DOC")) {
				doc.save(path5 + ".doc", SaveFormat.DOC);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("DOCX")) {

				doc.save(path5 + ".docx", SaveFormat.DOCX);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}

			} else if (filetype.equalsIgnoreCase("DOCM")) {

				doc.save(path5 + ".docm", SaveFormat.DOCM);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}

			} else if (filetype.equalsIgnoreCase("TIFF")) {

				doc.save(path5 + ".tiff", SaveFormat.TIFF);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}

			} else if (filetype.equalsIgnoreCase("TXT")) {

				doc.save(path5 + ".txt", SaveFormat.TEXT);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}

			} else if (filetype.equalsIgnoreCase("BMP")) {
				doc.save(path5 + ".bmp", SaveFormat.BMP);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("GIF")) {
				doc.save(path5 + ".gif", SaveFormat.GIF);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("JPG")) {

				doc.save(path5 + ".jpg", SaveFormat.JPEG);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("PNG")) {
				doc.save(path5 + ".png", SaveFormat.PNG);

				MapiAttachmentCollection attachments = message.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : message.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			}
		} catch (Exception e) {

			e.printStackTrace();
		}

		k = Main_Frame.count_destination++;

	}

}
