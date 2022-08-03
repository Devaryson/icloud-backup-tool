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
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;
import com.aspose.email.MapiMessage;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.facades.PdfContentEditor;
import com.aspose.words.Document;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;

public class MailWord implements Runnable {
	protected Main_Frame mf;
	protected String filetype = "";
	protected String destination_path = "";
	protected String path = "";
	protected String temppath = "";

	protected MailMessage message = null;
	protected Date reciveddate = null;
	protected Boolean Filter = null;
	protected Date startdate = null;
	protected Date enddate = null;
	long k;

	public MailWord(Main_Frame mf, String filetype, String destination_path, String path, String temppath,
			MailMessage message, Date reciveddate, boolean Filter, Date startdate, Date enddate) {
		this.mf = mf;
		this.filetype = filetype;
		this.destination_path = destination_path;
		this.path = path;
		this.temppath = temppath;
		this.message = message;
		this.Filter = Filter;
		this.startdate = startdate;
		this.enddate = enddate;

		this.reciveddate = reciveddate;

	}

	public void run() {
		mailword(message, reciveddate);
	}

	void mailword(MailMessage message, Date reciveddate) {

		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			// String input = mf.duplicacymapi(MapiMessage.fromMailMessage(message));
			String input = mf.duplicacymail(message);

			if (!mf.listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				mf.listduplicacy.add(input);

				if (Filter) {
					if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
						Mailmessage_word(message);

					} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
						Mailmessage_word(message);

					}
				} else {
					Mailmessage_word(message);

				}
			}
		} else {
			if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(startdate) && reciveddate.before(enddate)) {
					Mailmessage_word(message);

				} else if (reciveddate.equals(startdate) || reciveddate.equals(enddate)) {
					Mailmessage_word(message);

				}
			} else {
				Mailmessage_word(message);

			}
		}

	}

	@SuppressWarnings("unchecked")
	public void Mailmessage_word(MailMessage message) {
		String naming_convention = Main_Frame.getRidOfIllegalFileNameCharacters(mf.namingconventionmail(message));
		String path5 = destination_path + File.separator + path + File.separator + naming_convention + "_"
				+ Main_Frame.count_destination;
		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();

		message.save(emlStream, SaveOptions.getDefaultMhtml());
		LoadOptions lo = new LoadOptions();
		lo.setLoadFormat(LoadFormat.MHTML);
		try {
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			if (filetype.equalsIgnoreCase("PDF")) {
				path5 = path5.replaceAll("\\p{C}", "") + ".pdf";
				try {
					doc.save(path5, SaveFormat.PDF);
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				if (message.getAttachments().size() > 0) {

					MapiMessage msg = MapiMessage.fromMailMessage(message);
					PdfContentEditor editor = new PdfContentEditor();
					editor.bindPdf(path5);

					for (MapiAttachment attachment : msg.getAttachments()) {

						if (!(attachment.getExtension() == null)) {

							if (mf.chckbxSavePdfAttachment.isSelected()) {
								File f = new File(destination_path + File.separator + path + File.separator
										+ "Attachment" + File.separator + naming_convention	+ Main_Frame.count_destination);
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

								attachment.save(eml);
								editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
										attachment.getDisplayName(), "");
							}
						}

					}

					editor.save(path5);
					System.out.println("Embaded PDF file is created");

				}

				System.out.println("PDF Created!");

				System.out.print("save");

				// MapiMessage msg = MapiMessage.fromMailMessage(message);
//			System.out.println(message.getDate());
//			System.out.println(msg.getDeliveryTime());
//
//			try {
//
//				HtmlSaveOptions optionshtml = SaveOptions.getDefaultHtml();
//				optionshtml.setCheckBodyContentEncoding(true);
//				optionshtml.setEmbedResources(true);
//
//				optionshtml.setHtmlFormatOptions(HtmlFormatOptions.WriteCompleteBccEmailAddress
//						| HtmlFormatOptions.WriteHeader | HtmlFormatOptions.WriteCompleteCcEmailAddress
//						| HtmlFormatOptions.WriteCompleteFromEmailAddress
//						| HtmlFormatOptions.WriteCompleteToEmailAddress);
//
//				MapiAttachmentCollection attachments = msg.getAttachments();
//				File temppathpdf = new File(temppath + File.separator + "arysontemp");
//				temppathpdf.mkdirs();
//
//				Date c = message.getDate();
//				String date = c.toString();
//				String[] datearr = date.split("\\s");
//				String[] time = datearr[3].split(":");
//
//				Calendar call = Calendar.getInstance();
//				call.setTime(c);
//				call.set(Calendar.HOUR_OF_DAY, Integer.parseInt(time[0]));
//				call.set(Calendar.MINUTE, Integer.parseInt(time[1]));
//				call.set(Calendar.SECOND, Integer.parseInt(time[2]));
//				call.set(Calendar.MILLISECOND, 0);
//				call.setTimeZone(TimeZone.getDefault());
//				SimpleDateFormat sdf = new SimpleDateFormat("MMM dd HH:mm:ss z yyyy");
//			
//				String strDate = sdf.format(call.getTime());
//				Date c1 = null;
//
//				try {
//					c1 = sdf.parse(strDate);
//				} catch (ParseException e1) {
//
//					e1.printStackTrace();
//				}
//				message.setDate(c1);
//
//				Calendar cal = new GregorianCalendar();
//				TimeZone zone1 = cal.getTimeZone();
//				TimeZone zone = TimeZone.getTimeZone(zone1.getID());
//				int ms = zone.getOffset(Calendar.ZONE_OFFSET);
//				message.setTimeZoneOffset(ms);
//
//				message.save(temppath + File.separator + "1.html", optionshtml);
//
//				File url = new File(temppath + File.separator + "1.html");
//				ConverterProperties properties = new ConverterProperties();
//
//				properties.setFontProvider(new DefaultFontProvider(true, true, true));
//
//				String spath = "";
//				
//				if (mf.chckbxSavePdfAttachment.isSelected()) {
//					File fi=new File(path5);
//					fi.mkdirs();
//					path5=path5+File.separator+naming_convention;
//					
//				}
//
//				try {
//
//					HtmlConverter.convertToPdf(url, new File(path5.replaceAll("\\p{C}", "-") + ".pdf"), properties);
//					
//				} catch (Exception e) {
//					HtmlConverter.convertToPdf(url, new File(path5.replaceAll("\\p{C}", "-") + ".pdf"), properties);
//					
//				}
//
//				if (attachments.size() > 0) {
//					spath = path5.replaceAll("\\p{C}", "-") + ".pdf";
//
//					PDDocument doc = PDDocument.load(new File(spath));
//					List<PDEmbeddedFilesNameTreeNode> kids = new ArrayList<PDEmbeddedFilesNameTreeNode>();
//				
//
//					PDEmbeddedFilesNameTreeNode efTree = new PDEmbeddedFilesNameTreeNode();
//
//					for (MapiAttachment attachment : msg.getAttachments()) {
//
//						if (!(attachment.getExtension() == null)) {
//
//							if (mf.chckbxSavePdfAttachment.isSelected()) {
//								File f = new File(destination_path + File.separator + path + File.separator
//										+ "Attachment" + File.separator + naming_convention);
//								f.mkdirs();
//
//								String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
//								byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//								String str = new String(bytes, StandardCharsets.US_ASCII);
//								System.out.println(str);
//
//								attachment.save(f.getAbsolutePath() + File.separator
//										+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//							} else {
//
//								String s = attachment.getDisplayName();
//								byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
//								String str = new String(bytes, StandardCharsets.US_ASCII);
//								System.out.println(str);
//
//								attachment.save(temppath + File.separator + "arysontemp" + File.separator
//										+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								File att = new File(temppath + File.separator + "arysontemp" + File.separator
//										+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
//
//								
//
//								PDComplexFileSpecification fs = new PDComplexFileSpecification();
//								fs.setFile(att.getName());
//								byte[] data = Files.readAllBytes(att.toPath());
//								ByteArrayInputStream fakeFile = new ByteArrayInputStream(data);
//								PDEmbeddedFile ef = new PDEmbeddedFile(doc, fakeFile);
//
//								ef.setSize(data.length);
//								ef.setCreationDate(Calendar.getInstance());
//								fs.setEmbeddedFile(ef);
//
//								PDEmbeddedFilesNameTreeNode treeNode = new PDEmbeddedFilesNameTreeNode();
//								treeNode.setNames(Collections.singletonMap(attachment.getDisplayName(), fs));
//
//								kids.add(treeNode);
//
//								att.delete();
//							}
//						}
//
//					}
//
//					efTree.setKids(kids);
//
//					PDDocumentNameDictionary names1 = new PDDocumentNameDictionary(doc.getDocumentCatalog());
//					names1.setEmbeddedFiles(efTree);
//					doc.getDocumentCatalog().setNames(names1);
//					doc.save(new File(spath));
//					doc.close();
//					System.out.println("Embaded PDF file is created");
//				}
//
//				System.out.println("PDF Created!");
//
//				System.out.print("save");
//
//				new File(mf.filetemp).delete();
//				temppathpdf.delete();
//
//			} catch (Exception e) {
//				e.printStackTrace();
//			}
//		
			} else if (filetype.equalsIgnoreCase("DOC")) {
				doc.save(path5 + ".doc", SaveFormat.DOC);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("XPS")) {
				doc.save(path5 + ".xps", SaveFormat.XPS);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("DOCX")) {
				doc.save(path5 + ".docx", SaveFormat.DOCX);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("DOCM")) {
				doc.save(path5 + ".docm", SaveFormat.DOCM);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("Json")) {

				try {
					JSONObject employeeDetails = new JSONObject();
					employeeDetails.put("Subject", message.getSubject());
					employeeDetails.put("From", message.getFrom().toString());

					if (message.getSender() != null) {
						employeeDetails.put("Sender", message.getSender().toString());
					} else {
						employeeDetails.put("Sender", "");
					}

					employeeDetails.put("Date", message.getDate().toString());
					employeeDetails.put("Bcc", message.getBcc().toString());
					employeeDetails.put("Cc", message.getCC().toString());
					employeeDetails.put("Body", message.getBody().toString());
					if (new File(path5).isFile()) {
						new File(path5 + Main_Frame.count_destination).mkdirs();
						path5 = path5 + Main_Frame.count_destination;
					} else {
						new File(path5).mkdirs();

					}

					for (Attachment attachment : message.getAttachments()) {
						new File(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + naming_convention).mkdirs();
						employeeDetails.put("Attachment", attachment.getName().toString());
						attachment.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + naming_convention + File.separator
								+ attachment.getName().toString());
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
					System.out.println(employeeObject.toJSONString());
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("TIFF")) {
				doc.save(path5 + ".tiff", SaveFormat.TIFF);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("TXT")) {
				doc.save(path5 + ".txt", SaveFormat.TEXT);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("BMP")) {
				doc.save(path5 + ".bmp", SaveFormat.BMP);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("GIF")) {
				doc.save(path5 + ".gif", SaveFormat.GIF);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("JPG")) {
				doc.save(path5 + ".jpeg", SaveFormat.JPEG);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			} else if (filetype.equalsIgnoreCase("PNG")) {
				doc.save(path5 + ".png", SaveFormat.PNG);
				MapiMessage msg = MapiMessage.fromMailMessage(message);
				MapiAttachmentCollection attachments = msg.getAttachments();

				if (attachments.size() > 0) {
					for (MapiAttachment attachment : msg.getAttachments()) {
						attachment.save(path5 + File.separator + attachment.getDisplayName());

					}
				}
			}

		} catch (Exception e) {

			e.printStackTrace();
		}

		k = Main_Frame.count_destination++;
		message.dispose();

	}

}
