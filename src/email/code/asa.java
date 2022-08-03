package email.code;

import com.aspose.email.EmailClient;
import com.aspose.email.IConnection;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapFolderInfoCollection;
import com.aspose.email.SecurityOptions;

public class asa {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		ImapClient	clientforimap_input = new ImapClient("imap.mail.yahoo.com", 993, "aryson.test@yahoo.com", "bykijhfzvfgajchm");

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		 IConnection iconnforimap_input = clientforimap_input.createConnection();
		// clientforimap_input.deleteFolder(iconnforimap_input, "Trash");
	ImapFolderInfoCollection	asdf= clientforimap_input.listFolders();
	for(int  i=0;i<asdf.size();i++) {
		if(asdf.get_Item(i).getName()=="delete Items") {
			try {
				System.out.println(asdf.get_Item(i).getName());
				//clientforimap_input.del
				//clientforimap_input.deleteFolder(asdf.get_Item(i).getName());
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	//	 clientforimap_input.deleteFolder("Deleted Items");
	}

}
