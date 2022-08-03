package email.design;

import javax.swing.tree.DefaultMutableTreeNode;

@SuppressWarnings("serial")
public class CustomTreeNode extends DefaultMutableTreeNode {


	public String filepath;
	public CustomTreeNode(String folderName) {
		super(folderName);
	}
}

