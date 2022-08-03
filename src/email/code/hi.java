package email.code;

import com.aspose.email.FolderInfo;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiMessage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.StandardIpmFolder;

public class hi {
public static void main(String[] args) {
	PersonalStorage pst=PersonalStorage.create("E:\\New folder (2)\\ai.pst", 0);
	FolderInfo info=pst.createPredefinedFolder("Task",StandardIpmFolder.Tasks);
	MailMessage msg=MailMessage.load("C:\\Users\\NEW\\Desktop\\Thu Jul 15 12 04 46 IST 2021\\DuplicateOTest\\Top of Outlook data file\\Tasks\\987798879.eml");
	info.addMessage(MapiMessage.fromMailMessage(msg));
	 msg=MailMessage.load("C:\\Users\\NEW\\Desktop\\Thu Jul 15 12 04 46 IST 2021\\DuplicateOTest\\Top of Outlook data file\\Tasks\\dffdddfdf.eml");
	 info.addMessage(MapiMessage.fromMailMessage(msg));
	 System.out.println("Hi");
}
}
