package email.code;

import java.util.Date;

import com.aspose.email.MailMessage;

public class Mailimap implements Runnable {
	Main_Frame mf;
	String path = "";

	Date reciveddate = null;
	MailMessage message = null;
	String Messageid = "";

	public Mailimap(Main_Frame mf, String path, MailMessage message, Date reciveddate) {
		this.mf = mf;
		this.path = path;
		this.message = message;
		this.reciveddate = reciveddate;

	}

	public void run() {
		try {
			Messageid = mailimap(message, reciveddate, path);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	String mailimap(MailMessage message, Date reciveddate, String path) throws Exception {
		String Messageid = "";
		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!mf.listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				mf.listduplicacy.add(input);

				if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mf.mailfilterstartdate) && reciveddate.before(mf.mailfilterenddate)) {
						Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
								message);
						Main_Frame.foldermessagecount++;
						Main_Frame.count_destination++;

					} else if (reciveddate.equals(mf.mailfilterstartdate) || reciveddate.equals(mf.mailfilterenddate)) {
						Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
								message);
						Main_Frame.foldermessagecount++;
						Main_Frame.count_destination++;

					}
				} else {

					Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
							message);
					Main_Frame.foldermessagecount++;
					Main_Frame.count_destination++;
				}
			} else {
				System.out.println(" duplicate message");
				System.out.println(input);

			}
		} else {
			if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mf.mailfilterstartdate) && reciveddate.before(mf.mailfilterenddate)) {
					Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
							message);
					Main_Frame.foldermessagecount++;
					Main_Frame.count_destination++;

				} else if (reciveddate.equals(mf.mailfilterstartdate) || reciveddate.equals(mf.mailfilterenddate)) {

					Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
							message);
					Main_Frame.foldermessagecount++;
					Main_Frame.count_destination++;
				}
			} else {
				Messageid = Main_Frame.clientforimap_output.appendMessage(Main_Frame.iconnforimap_output, path,
						message);

				Main_Frame.count_destination++;

			}
		}
		return Messageid;
	}
}
