package email.code;

import java.io.File;

import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.PersonalStorage;

public class hoh {
	
	public static void main(String[] args) {
		
		File f=new File("C:\\Users\\NEW\\Desktop\\hih");
		File [] f1=f.listFiles();
		PersonalStorage pst= PersonalStorage.create("C:\\\\Users\\\\NEW\\\\Desktop\\3.pst",0);
		
		
		for (int i = 0; i < f1.length; i++) {
			try {
				File dd=f1[i];
				    PersonalStorage ps=PersonalStorage.fromFile(dd.getAbsolutePath());
				    FolderInfo mergef=pst.getRootFolder().addSubFolder(dd.getName());
				    
				   String  path=dd.getName()+File.separator+ps.getRootFolder().getDisplayName();
				         mergef.addSubFolder(path,true);
				         
				         System.out.println(path+File.separator+ps.getRootFolder().getDisplayName());
				         
				        FolderInfoCollection folderInfoCollection= ps.getRootFolder().getSubFolders();
				         for (int j = 0; j < folderInfoCollection.size(); j++) {
						FolderInfo dddd=	folderInfoCollection.get(j);
						
						
						mergef.addSubFolder(path+File.separator+dddd.getDisplayName(),true);
						
						
						if(dddd.hasSubFolders()) {
							
							subfolder(dddd,mergef,path+File.separator+dddd.getDisplayName());
							
						}
				        	 
				        	 
				        	 
						}
			} catch (Exception e) {
				
				e.printStackTrace();
			}
		             
		        
		        
		        
			
			
		}
		
		
		
		
	}
	
	static void subfolder(FolderInfo mainfolder,FolderInfo mergef,String path) {
		
	FolderInfoCollection folderInfoCollection=	mainfolder.getSubFolders();
		try {
			for (int i = 0; i < folderInfoCollection.size(); i++) {
				FolderInfo subfolderpst=folderInfoCollection.get(i);
				path=path+File.separator+subfolderpst.getDisplayName();
				mergef.addSubFolder(path,true);
				
				System.out.println(path);
				if(subfolderpst.hasSubFolders()) {
					
					subfolder(subfolderpst,mergef,path);
					
					
				}
				
				
				
				
				
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
		
		
		
	};
	

}
