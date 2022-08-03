package email.code;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.EventQueue;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import java.util.Calendar;
import java.util.Date;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiMessage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.chilkatsoft.CkEmail;
import com.chilkatsoft.CkEmailBundle;
import com.chilkatsoft.CkGlobal;
import com.chilkatsoft.CkImap;
import com.chilkatsoft.CkMailboxes;
import com.chilkatsoft.CkMessageSet;

import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import java.awt.Font;
import javax.swing.JTextField;
import javax.swing.JLabel;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.CardLayout;
import javax.swing.JTree;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.JScrollPane;
import javax.swing.JComboBox;

public class teste  {

	/**
	 * 
	 */
	private static final long serialVersionUID = -5987755994305099258L;
	private JPanel contentPane;
	private JTextField textField;
	private JTextField textField_1;
	CheckboxTree tree;
	CkImap imap;
	String wildcardedMailbox = "*";
	String refName = "";
	DefaultTreeModel model;
	DefaultMutableTreeNode root;
	private List<String> listFolderinfostring=new ArrayList<String>();
	JPanel Cardlayout;
	private List<String> folderlist=new ArrayList<String>();
	JComboBox comboBox ;
	boolean success;
	JLabel lblNewLabel_3_2_1;
	JLabel lblNewLabel_2;
	long countdest;
	private JButton btnNewButton_2;
	JLabel lblNewLabel_3_1;
	JLabel lblNewLabel_3_2;
	JLabel lblNewLabel_3;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					System.out.println(File.separator);
					System.out.println("\\");
					System.out.println("\\\\");
					System.out.println("/");
				} catch (Exception e1) {

				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
}