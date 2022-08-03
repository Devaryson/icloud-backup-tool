package email.code;
import java.util.Date;
import com.aspose.email.MailMessage;

public class Mailexchange implements Runnable {
	Main_Frame mf;
	String Folderuri = "";
	
	Date reciveddate = null;
	MailMessage message = null;
	String Messageid = "";

	public Mailexchange(Main_Frame mf, String Folderuri, MailMessage message, Date reciveddate) {
		this.mf = mf;
		this.Folderuri = Folderuri;
		this.message = message;
		this.reciveddate = reciveddate;

	}

	public void run() {
		Messageid = mailexchange1(message, reciveddate, Folderuri);
	}

	String mailexchange1(MailMessage message, Date reciveddate, String Folderuri) {
		String Messageid = "";
		if (mf.chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!mf.listduplicacy.contains(input)) {

				mf.listduplicacy.add(input);

				if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mf.mailfilterstartdate) && reciveddate.before(mf.mailfilterenddate)) {
						Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

						Main_Frame.count_destination++;

					} else if (reciveddate.equals(mf.mailfilterstartdate) || reciveddate.equals(mf.mailfilterenddate)) {
						Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

						Main_Frame.count_destination++;

					}
				} else {

					Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

					Main_Frame.count_destination++;
				}
			} else {

			}
		} else {
			if (Main_Frame.chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mf.mailfilterstartdate) && reciveddate.before(mf.mailfilterenddate)) {
					Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

					Main_Frame.count_destination++;

				} else if (reciveddate.equals(mf.mailfilterstartdate) || reciveddate.equals(mf.mailfilterenddate)) {

					Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

					Main_Frame.count_destination++;
				}
			} else {
				Messageid = Main_Frame.clientforexchange_output.appendMessage(Folderuri, message);

				Main_Frame.count_destination++;

			}
		}
		return Messageid;
	}
}
