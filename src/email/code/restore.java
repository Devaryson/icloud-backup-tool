package email.code;

import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.net.InetAddress;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.UnknownHostException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.Executors;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.ButtonGroup;
import javax.swing.DefaultCellEditor;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import javax.swing.text.html.HTMLEditorKit;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;

import com.aspose.email.Appointment;
import com.aspose.email.AppointmentLoadOptions;
import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.Contact;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.EWSClient;
import com.aspose.email.EmailAddress;
import com.aspose.email.EmailAddressList;
import com.aspose.email.EmailClient;
import com.aspose.email.EmlSaveOptions;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeMailboxInfo;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapMessageInfo;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;

import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MboxLoadOptions;
import com.aspose.email.MboxrdStorageReader;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.OlmFolder;
import com.aspose.email.OlmStorage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SecurityOptions;
import com.aspose.email.system.io.FileAccess;
import com.aspose.email.system.io.FileMode;
import com.aspose.email.system.io.FileStream;
import com.opencsv.CSVWriter;
import com.toedter.calendar.JDateChooser;

import au.com.bytecode.opencsv.CSVReader;
import email.activation.OnlineActivation;
import email.activation.Starting_Frame;
import email.design.HeaderRenderer2;
import email.design.ImageRenderer;
import email.design.ImageRenderer1;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.DefaultCheckboxTreeCellRenderer;
import javax.swing.border.TitledBorder;
import javax.swing.UIManager;
import javax.swing.WindowConstants;

@SuppressWarnings("deprecation")
public class restore extends JFrame {

	private final ButtonGroup buttonGroup = new ButtonGroup();
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private static JTable table;
	static long count_destination;
	static JFileChooser jFileChooser;
	static File[] files;
	JPanel panel_7;
	JComboBox<String> comboBox_1;
	static JPanel Cardlayout;
	loadingThreadclassformailbox obThmailbox;
	JComboBox<String> uservaluescomboBox;
	JCheckBox chckbxrestoretodefault;
	LoadingThreadclass obTh;
	TableColumnModel csvcolumnModel2;
	String[] usernameverification;
	JTable table_userid;
	JPanel panel_Contact;
	JPanel panel_Callendar;
	Boolean contatcheck;
	Boolean calendarcheck;
	JLabel lblFullName;
	JLabel lblEMailAd;
	JLabel lblPhoneNo;
	JLabel label_contactcompany;
	JLabel llabel_contactfullname;
	JLabel lblCompany;
	JLabel label_contactemail;
	JLabel label_contactfullname;
	JLabel label_contactphonenumber;
	JLabel label_contacticon;
	JLabel label_calendarstartdate;
	JLabel lblEndDate;
	JLabel label_Calendarsubject;
	JLabel label_Calendaricon;
	JTextArea textArea_contact;
	static DefaultTableModel modeli;
	static JComboBox<String> comboBox_FiletypeChooser;
	static String fileoptionm = "MICROSOFT OUTLOOK (.pst)";
	File f;
	JCheckBox chckbxRemoveDuplicacy;
	JComboBox<String> comboBox;
	File filetem;
	JButton btn_signout_p3;
	Calendar cal;
	JPanel panel_progress;
	JPanel panel_taskfilter;
	long count_eml_msg_emlx = 0;
	int portnofiletype;
	File file;
	static int countforfile = 0;
	static PersonalStorage pst;
	PersonalStorage ost;
	Main_Frame mf;
	List<String> pstfolderlist;
	ArrayList<String> pstfolderlist2;
	static Date mailfilterstartdate;
	static Date mailfilterenddate;
	static Date Calenderfilterstartdate;
	static Date Calenderfilterenddate;
	static Date taskfilterenddate;
	static Date taskfilterstartdate;
	JButton btn_previous_p2;
	JButton btn_remove;
	JLabel lbl_header;
	Boolean output = false;
	JLabel lbl_Domain;
	HashMap<String, List<String>> hm;

	String fa = "";
	JLabel lbl_connecting_p3;
	JLabel lbl_progressreport;
	FolderInfo info = new FolderInfo();
	FolderInfo folderInfo;
	static DefaultMutableTreeNode root;
	static DefaultTreeModel model;
	DefaultMutableTreeNode mainnode;
	static DefaultMutableTreeNode lastNode;
	static DefaultTableModel mode;
	JButton btn_select_folder;
	JButton btn_previous_p3;
	JButton btn_Destination;
	JButton btn_cancel;
	JButton btnStop;
	private JButton btn_converter_1;
	JButton btn_Next;
	JButton btnDowloadReport;
	private JButton btnDowloadReport_1;
	String domain_p3;
	String username_p3;
	String password_p3;
	String filetype = "Office 365";
	String calendertime;
	String parentfolder;
	String subfolderfile;
	String reportpath;
	String filepath;

	String Status = "Completed";
	String fname;
	String filetemp = "";
	String foldername = "";
	String foldername2 = "";
	String foldername3 = "";
	String foldername4 = "";
	String path = "";
	String path2;
	String path3 = "";
	String parentname;
	String path1 = "";
	String s = "";
	String parent = "";
	String[] filesfin;
	String Folder;
	String Folderuri;
	String destination_path;
	static String messageboxtitle;
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	String buyurl;
	String infourl;
	String helpurl;
	MboxrdStorageWriter wr;
	OlmStorage storage;
	int filesno = 1;
	int folder = 0;
	CSVWriter writer;
	boolean connectioncheck = true;
	JPanel panel_3;
	JPanel panel_3_;
	JPanel panel_3_1;
	JPanel panel_3_1_1;
	JLabel lblLive_Chat_p3;
	JPanel panel_3_2;
	JPanel panel_4;
	JPanel attachmenttable;
	JPanel innercardlayout;
	JPanel panel_mailboxid;
	JPanel panel_3_1_2;
	JLabel label_11;
	JLabel lblMakeSureYou;
	JLabel lblEnableImap_p3;
	JLabel lblTurnOffTwo_p3;
	JLabel lblAllowLessSecure_p3;
	JLabel label_10;
	JLabel lblLoadingPleaseWait;
	JTextField textField_username_p3;
	JTextField tf_Destination_Location;
	JPasswordField passwordField_p3;

	Thread th;
	static IEWSClient clientforexchange_output;
	static ImapClient clientforimap_output;
	static IConnection iconnforimap_output;
	JProgressBar progressBar_message_p3;
	JTable table_fileConvertionreport_panel4;
	JComboBox<String> comboBox_fileDestination_type;
	JButton btn_next_pane2;
	Calendar Cal;
	static Boolean demo = true;
	Boolean btnfile = false;
	Boolean btnfolder = false;
	Boolean stop = false;
	Boolean pstCalenderfoldercheck = true;
	Boolean pstContactfoldercheck = true;
	Boolean psttaskfoldercheck = true;
	Boolean foldercheck = true;
	Boolean checkconvertagain = false;
	Boolean checkdestination = true;
	Boolean Stoppreview = false;
	int ret;
	static CheckboxTree tree;
	OlmFolder folderi;
	long foldermessagecount;
	FolderInfo info1;
	JDateChooser dateChooser_calender_start;
	JDateChooser dateChooser_mail_fromdate;
	JDateChooser dateChooser_task_end_date;
	JDateChooser dateChooser_task_start_date;
	JDateChooser dateChooser_mail_tilldate;
	JDateChooser dateChooser_calendar_end;
	JButton btnremove_all;
	Boolean checky = true;
	JCheckBox chckbx_Mail_Filter;
	JCheckBox chckbx_calender_box;

	private JTable table_fileinformation;

	private JEditorPane editorPane;
	private JScrollPane scrollPane_3;
	private JTable table_1;
	private JPanel panel_3_1_2_1;
	private JTextField textField_domain_name_p3;
	Set<File> hashset = new LinkedHashSet<File>();
	private JLabel lblNewLabel_4;
	private JScrollPane scrollPane_4;
	JCheckBox chckbxCustomFolderName;
	JCheckBox task_box;

	private List<String> listduplicacy = new ArrayList<String>();
	String logpathm = "";
	String temppathm = "";
	private JLabel Progressbar;
	private JLabel lblPortNo;
	private JTextField textField;
	private JPanel panel;
	private JPanel panel_5;
	private JLabel lblnamingconvention;
	private JPanel panel_6;
	private JTextField textField_customfolder;
	private JLabel label_Calendarenddate;
	private JLabel lblNotes;
	private JTextField textField_importcsvpath;
	private JLabel lblNewLabel_5;
	private JPanel panel_fileselect;
	private JTable table_fileselect;
	private JPanel panel_mailfolder;
	private JTable table_2;
	static List<DefaultMutableTreeNode> lists = new ArrayList<DefaultMutableTreeNode>();
	static List<String> listst = new ArrayList<String>();
	private List<MailMessage> listmail = new ArrayList<MailMessage>();
	private List<MapiMessage> listmapi = new ArrayList<MapiMessage>();
	private List<ImapMessageInfo> listImapmesinfo = new ArrayList<ImapMessageInfo>();
	private List<ExchangeMessageInfo> listExchangemesingo = new ArrayList<ExchangeMessageInfo>();
	private List<MessageInfo> listPSTOSTgemesingo = new ArrayList<MessageInfo>();
	private List<String> userList = new ArrayList<String>();
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
	double SPACE_KB = 1024;
	double SPACE_MB = 1024 * SPACE_KB;
	double SPACE_GB = 1024 * SPACE_MB;
	double SPACE_TB = 1024 * SPACE_GB;
	String version;
	static Boolean fileoptioncheck;
	private JLabel lblTotalMessageCount;
	private JLabel label_12;

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public restore(JFrame parent1, Boolean demo, String messageboxtitle) {
		super();
		mf = (Main_Frame) parent1;
		restore.demo = demo;
		restore.messageboxtitle = messageboxtitle;
		fileoptionm = Main_Frame.fileoption;
		if (fileoptionm.equalsIgnoreCase("Mbox to Office 365")) {
			fileoptionm = "Mbox";
		} else {
			fileoptionm = "MICROSOFT OUTLOOK (.pst)";
		}

		temppathm = mf.temppath;
		logpathm = mf.logpath;
		calendertime = Main_Frame.calendertime;
		buyurl = mf.buyurl;
		version = mf.version;
		infourl = mf.infourl;
		helpurl = mf.helpurl;

		addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent arg0) {
				String warn = "Do you want to close the application?";
				int ans = JOptionPane.showConfirmDialog(restore.this, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					if (demo) {

						openBrowser(All_Data.infourl);

					}
					System.exit(0);
				}
			}
		});

		setIconImage(Toolkit.getDefaultToolkit().getImage(Main_Frame.class.getResource("/128x128.png")));

		setTitle(messageboxtitle);
		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		setLocationRelativeTo(null);
		setResizable(false);
		setBounds(100, 100, 1081, 718);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		Cardlayout = new JPanel();
		Cardlayout.setBackground(Color.LIGHT_GRAY);
		Cardlayout.setBounds(0, 71, 1075, 618);
		contentPane.add(Cardlayout);
		Cardlayout.setLayout(new CardLayout(0, 0));

		JPopupMenu menu = mf.menu;

		JPanel panel_1 = new JPanel();
		panel_1.setBackground(Color.WHITE);
		Cardlayout.add(panel_1, "panel_1");
		panel_1.setLayout(null);

		JButton btnActivate = new JButton("");
		btnActivate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
			}
		});

		btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
		btnActivate.setRolloverEnabled(false);
		btnActivate.setRequestFocusEnabled(false);
		btnActivate.setOpaque(false);
		btnActivate.setFocusable(false);
		btnActivate.setFocusTraversalKeysEnabled(false);
		btnActivate.setFocusPainted(false);
		btnActivate.setDefaultCapable(false);
		btnActivate.setContentAreaFilled(false);
		btnActivate.setBorderPainted(false);
		if (demo) {
			btnActivate.setVisible(true);
		} else {
			btnActivate.setVisible(false);

		}
		btnActivate.setToolTipText("Click here to activate the software");
		btnActivate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				File licFileon = null;
				if (System.getProperty("os.name").toLowerCase().contains("windows")) {
					licFileon = new File(System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator
							+ "licenseOnline.lic");

//					licFileonoff = new File(
//							System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator + "license.lic");

				} else {
					licFileon = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
							+ "Application Support" + File.separator + messageboxtitle + File.separator
							+ "licenseOnline.lic");
//					licFileonoff = new File(
//							System.getenv("APPDATA") + File.separator + messageboxtitle + File.separator + "license.lic");
				}
				boolean activatefromdemo = true;
				new Starting_Frame();
				OnlineActivation mf = new OnlineActivation(Starting_Frame.mf, licFileon, activatefromdemo);
				mf.setLocationRelativeTo(null);
				mf.setVisible(true);
				setEnabled(false);
				mf.btnBack.setVisible(false);

				mf.addWindowListener(new WindowAdapter() {
					@Override
					public void windowClosing(WindowEvent arg0) {
						String warn = "Do you want to close?";
						int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
								JOptionPane.QUESTION_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
						if (ans == JOptionPane.YES_OPTION) {
							setEnabled(true);
							mf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
						} else {
							mf.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);

						}
					}
				});

			}
		});

		btnActivate.setBounds(972, 17, 32, 40);
		contentPane.add(btnActivate);

		JButton btn_info = new JButton("");
		btn_info.setToolTipText("Click here For Software Information.");
		btn_info.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				AboutDialog ab;
				if (demo) {
					ab = new AboutDialog(mf, true, "demo");

				} else {
					ab = new AboutDialog(mf, true, "full");
				}
				ab.setLocationRelativeTo(mf);
				ab.setVisible(true);

			}
		});
		btn_info.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
			}
		});

		JButton btn_buy = new JButton("");
		btn_buy.setToolTipText("Click here to buy the Software.");
		btn_buy.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openBrowser(buyurl);

			}
		});
		if (demo) {
			btn_buy.setVisible(true);
		} else {
			btn_buy.setVisible(false);
		}

		JButton btn_help = new JButton("");
		btn_help.setToolTipText("Click here For Software Guide.");
		btn_help.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(helpurl);
			}
		});
		btn_help.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));
			}
		});

		btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));

		btn_help.setRolloverEnabled(false);
		btn_help.setRequestFocusEnabled(false);
		btn_help.setOpaque(false);
		btn_help.setFocusable(false);
		btn_help.setFocusTraversalKeysEnabled(false);
		btn_help.setFocusPainted(false);
		btn_help.setDefaultCapable(false);
		btn_help.setContentAreaFilled(false);
		btn_help.setBorderPainted(false);
		btn_help.setBounds(998, 17, 41, 41);
		contentPane.add(btn_help);
		btn_buy.setBounds(938, 17, 41, 41);
		contentPane.add(btn_buy);

		btn_buy.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));
			}
		});

		btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));

		btn_buy.setOpaque(false);
		btn_buy.setRolloverEnabled(false);
		btn_buy.setRequestFocusEnabled(false);
		btn_buy.setFocusTraversalKeysEnabled(false);
		btn_buy.setFocusable(false);
		btn_buy.setFocusPainted(false);
		btn_buy.setDefaultCapable(false);
		btn_buy.setContentAreaFilled(false);
		btn_buy.setBorderPainted(false);

		btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
		btn_info.setRolloverEnabled(false);
		btn_info.setRequestFocusEnabled(false);
		btn_info.setOpaque(false);
		btn_info.setFocusable(false);
		btn_info.setFocusTraversalKeysEnabled(false);
		btn_info.setFocusPainted(false);
		btn_info.setDefaultCapable(false);
		btn_info.setContentAreaFilled(false);
		btn_info.setBorderPainted(false);
		btn_info.setBounds(1025, 17, 47, 41);
		contentPane.add(btn_info);

		btn_Next = new JButton("");
		btn_Next.setToolTipText("Click here to Go Forward.");
		btn_Next.setVisible(false);
		btn_Next.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Next.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_Next.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
			}
		});

		btn_Next.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));

		btn_Next.setRolloverEnabled(false);
		btn_Next.setRequestFocusEnabled(false);
		btn_Next.setOpaque(false);
		btn_Next.setFocusable(false);
		btn_Next.setFocusTraversalKeysEnabled(false);
		btn_Next.setFocusPainted(false);
		btn_Next.setDefaultCapable(false);
		btn_Next.setContentAreaFilled(false);
		btn_Next.setBorderPainted(false);
		btn_Next.setBounds(933, 567, 123, 38);
		panel_1.add(btn_Next);
		btn_Next.setEnabled(false);
		btn_Next.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				lists.clear();
				listst.clear();
				Object[] s = new Object[table.getRowCount()];
				filesfin = new String[table.getRowCount()];

				for (int i = 0; i < table.getRowCount(); i++) {
					s[i] = table.getValueAt(i, 2);
				}

				for (int i = 0; i < s.length; i++) {

					filesfin[i] = (String) s[i];

				}

				mf.logger = mf.logFile();
				mf.logger.info("Start Time : " + calendertime + System.lineSeparator() + "File filetype  " + fileoptionm
						+ System.lineSeparator()
						+ "======================================================================");

				SwingWorker sw1 = new SwingWorker() {
					@Override
					protected Object doInBackground() {
						obTh = new LoadingThreadclass(mf);
						obTh.start();
						InetAddress addr;
						String hostName = "";
						try {
							addr = InetAddress.getLocalHost();
							hostName = addr.getHostName();

						} catch (UnknownHostException e1) {

							e1.printStackTrace();
						}

						model = (DefaultTreeModel) tree.getModel();

						root = new DefaultMutableTreeNode("<html><b>" + hostName);

						model.setRoot(root);

						mainnode = new DefaultMutableTreeNode("<html><b>" + fileoptionm);

						root.add(mainnode);

						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_2");

						for (int i = 0; i < filesfin.length; i++) {

							String filetype = table.getValueAt(i, 3).toString();
							filetype = filetype.replace("<html><b>", "");
							if (filetype.equalsIgnoreCase("file")) {

								filepath = filesfin[i].replace("<html><b>", "");
								path2 = filepath;

								if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									try {
										String extension = getFileExtension(new File(path2));
										if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
											if (extension.equalsIgnoreCase("pst")
													|| extension.equalsIgnoreCase("ost")) {

												try {
													readAnOST_PstFile();
												} catch (Exception e) {
													i--;
													mainnode.removeAllChildren();
													continue;
												}
											}

										} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

											if (extension.equalsIgnoreCase("ost")) {

												try {
													readAnOST_PstFile();
												} catch (Exception e) {
													i--;
													mainnode.removeAllChildren();
													continue;
												}
											}
										}

									} catch (Exception e1) {

										e1.printStackTrace();
									}

								} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
									try {
										readMboxFile();
									} catch (Exception e) {
										i--;
										mainnode.removeAllChildren();
										continue;
									}
								} else if (fileoptionm.equalsIgnoreCase("OLM files (.olm)")) {
									readolmFile();
								} else {
									String extension = getFileExtension(new File(path2));
									if (fileoptionm.equalsIgnoreCase("EML Files (.eml)")) {
										if (extension.equals("eml")) {

											readmailFile();

										}
									} else if (fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")) {
										if (extension.equalsIgnoreCase("emlx")) {

											readmailFile();

										}
									} else if (fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {
										if (extension.equalsIgnoreCase("msg")) {

											readmailFile();
										}
									}

								}
							} else {

								if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									try {
										read_PSTOST_folder(new File(filesfin[i].replace("<html><b>", "")));
									} catch (Exception e1) {

										e1.printStackTrace();
									}

								} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
									try {
										read_mbox_folder(new File(filesfin[i].replace("<html><b>", "")));
									} catch (Exception e1) {

										e1.printStackTrace();

									}

								} else if (fileoptionm.equalsIgnoreCase("OLM files (.olm)")) {
									try {
										read_olm_folder(new File(filesfin[i].replace("<html><b>", "")));
									} catch (Exception e1) {
										e1.printStackTrace();
									}

								} else {

									reademl_emlx_msg_folder(new File(filesfin[i].replace("<html><b>", "")));
								}

							}
						}
						return null;
					}

					@Override
					protected void done() {
//						
						Icon open = new ImageIcon(backup.class.getResource("/Open-folder-accept-icon.png"));
						Icon close = new ImageIcon(backup.class.getResource("/closed-folder-add-icon.png"));
						Icon Ram = new ImageIcon(backup.class.getResource("/leaf-icon.png"));
						DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree
								.getCellRenderer();
						render.setClosedIcon(close);
						render.setOpenIcon(open);
						render.setLeafIcon(Ram);

						tree.expandRow(0);
						tree.expandAll();
						obTh.close();
					}
				};

				sw1.execute();

			}

		});
		btn_Next.setFont(new Font("Tahoma", Font.BOLD, 12));

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(235, 0, 840, 552);
		panel_1.add(scrollPane);

		table = new JTable() {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};
		scrollPane.setViewportView(table);
		table.getTableHeader().setReorderingAllowed(false);
		table.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "S.No", " Name", "File Path", "File Type", "Size" }));
		table.getColumnModel().getColumn(0).setPreferredWidth(35);
		table.getColumnModel().getColumn(1).setPreferredWidth(57);
		table.getColumnModel().getColumn(2).setPreferredWidth(350);
		table.getColumnModel().getColumn(3).setPreferredWidth(20);
		table.getColumnModel().getColumn(4).setPreferredWidth(40);

		lbl_header = new JLabel("");
		lbl_header.setFont(new Font("Tahoma", Font.BOLD, 15));
		lbl_header.setBounds(392, 13, 390, 27);
		panel_1.add(lbl_header);

		JButton btn_previous = new JButton("");
		btn_previous.setToolTipText("Click here to Go Back.");
		btn_previous.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		btn_previous.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});
		btn_previous.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous.setRolloverEnabled(false);
		btn_previous.setRequestFocusEnabled(false);
		btn_previous.setOpaque(false);
		btn_previous.setFocusable(false);
		btn_previous.setFocusTraversalKeysEnabled(false);
		btn_previous.setFocusPainted(false);
		btn_previous.setDefaultCapable(false);
		btn_previous.setContentAreaFilled(false);
		btn_previous.setBorderPainted(false);
		if (All_Data.messageboxtitle.contains("office 365") || All_Data.messageboxtitle.contains("Outlook"))
			btn_previous.setVisible(false);
		btn_previous.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Main_Frame main_m = new Main_Frame(demo, Main_Frame.versiontype);
				main_m.setVisible(true);
				main_m.setLocationRelativeTo(null);
				main_m.setResizable(false);
				main_m.rdbtnSingleFile.setSelected(true);
				dispose();
			}
		});
		btn_previous.setFont(new Font("Tahoma", Font.BOLD, 12));
		btn_previous.setBounds(811, 567, 123, 38);
		panel_1.add(btn_previous);

		lblNewLabel_4 = new JLabel("");
		lblNewLabel_4.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		lblNewLabel_4.setBackground(Color.WHITE);
		lblNewLabel_4.setBounds(0, 550, 1075, 68);
		panel_1.add(lblNewLabel_4);

		JPanel panel_8 = new JPanel();
		panel_8.setBackground(Color.WHITE);
		panel_8.setBounds(0, 0, 234, 546);
		panel_1.add(panel_8);
		panel_8.setLayout(null);

		JButton btn_ChoseFile = new JButton("");
		btn_ChoseFile.setToolTipText("Click here to Select File(s).");
		btn_ChoseFile.setBounds(20, 31, 118, 39);
		panel_8.add(btn_ChoseFile);
		btn_ChoseFile.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_ChoseFile.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-files-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_ChoseFile.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-files-btn.png")));
			}
		});

		btn_ChoseFile.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-files-btn.png")));
		btn_ChoseFile.setRolloverEnabled(false);
		btn_ChoseFile.setRequestFocusEnabled(false);
		btn_ChoseFile.setOpaque(false);
		btn_ChoseFile.setFocusable(false);
		btn_ChoseFile.setFocusTraversalKeysEnabled(false);
		btn_ChoseFile.setFocusPainted(false);
		btn_ChoseFile.setDefaultCapable(false);
		btn_ChoseFile.setContentAreaFilled(false);
		btn_ChoseFile.setBorderPainted(false);
		btn_ChoseFile.setBackground(Color.WHITE);
		btn_ChoseFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				try {
					filter_file();
					btnfile = true;
					btn_Next.setEnabled(true);
					btn_Next.setVisible(true);
					btn_remove.setEnabled(true);
					btnremove_all.setEnabled(true);

					if (countforfile == 0) {

						btnfile = false;
						btn_Next.setEnabled(false);
						btn_Next.setVisible(false);
						btn_remove.setEnabled(false);
						btnremove_all.setEnabled(false);
					}
				} catch (Exception e) {

				}

			}
		});

		btn_ChoseFile.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_select_folder = new JButton("");
		btn_select_folder.setToolTipText("Click here to Select Folder(s).");
		btn_select_folder.setBounds(20, 146, 118, 39);
		panel_8.add(btn_select_folder);
		btn_select_folder.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_select_folder.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-folder-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_select_folder.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-folder-btn.png")));
			}
		});

		btn_select_folder.setIcon(new ImageIcon(Main_Frame.class.getResource("/select-folder-btn.png")));
		btn_select_folder.setOpaque(false);
		btn_select_folder.setFocusable(false);
		btn_select_folder.setFocusTraversalKeysEnabled(false);
		btn_select_folder.setFocusPainted(false);
		btn_select_folder.setDefaultCapable(false);
		btn_select_folder.setContentAreaFilled(false);
		btn_select_folder.setBorderPainted(false);
		btn_select_folder.setRolloverEnabled(false);
		btn_select_folder.setRequestFocusEnabled(false);
		btn_select_folder.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_remove = new JButton("");
		btn_remove.setToolTipText("Click here to Remove Selected File(s).");
		btn_remove.setBounds(20, 262, 123, 38);
		panel_8.add(btn_remove);
		btn_remove.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-btn.png")));
			}
		});

		btn_remove.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-btn.png")));
		btn_remove.setRolloverEnabled(false);
		btn_remove.setRequestFocusEnabled(false);
		btn_remove.setOpaque(false);
		btn_remove.setFocusable(false);
		btn_remove.setFocusTraversalKeysEnabled(false);
		btn_remove.setFocusPainted(false);
		btn_remove.setDefaultCapable(false);
		btn_remove.setContentAreaFilled(false);
		btn_remove.setBorderPainted(false);
		btn_remove.setEnabled(false);
		btn_remove.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				try {
					DefaultTableModel model = (DefaultTableModel) table.getModel();
					int selected = table.getSelectedRow();

					hashset.remove(new File(table.getValueAt(selected, 2).toString().replace("<html><b>", "")));

					model.removeRow(selected);
					countforfile--;
					filesno--;
					if (countforfile == 0) {

						btn_Next.setEnabled(false);
						btn_Next.setVisible(false);
						btn_remove.setEnabled(false);
						btnremove_all.setEnabled(false);
					}
				} catch (Exception e) {
					JOptionPane.showMessageDialog(mf, "Please select  file or Folder you want to remove",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
				}

			}
		});
		btn_remove.setFont(new Font("Tahoma", Font.BOLD, 12));

		btnremove_all = new JButton("");
		btnremove_all.setToolTipText("Click here to Remove All File(s).");
		btnremove_all.setBounds(20, 376, 118, 39);
		panel_8.add(btnremove_all);
		btnremove_all.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnremove_all.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-all-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnremove_all.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-all-btn.png")));
			}
		});

		btnremove_all.setIcon(new ImageIcon(Main_Frame.class.getResource("/remove-all-btn.png")));
		btnremove_all.setRolloverEnabled(false);
		btnremove_all.setRequestFocusEnabled(false);
		btnremove_all.setOpaque(false);
		btnremove_all.setFocusable(false);
		btnremove_all.setFocusTraversalKeysEnabled(false);
		btnremove_all.setFocusPainted(false);
		btnremove_all.setDefaultCapable(false);
		btnremove_all.setContentAreaFilled(false);
		btnremove_all.setBorderPainted(false);
		btnremove_all.setEnabled(false);
		btnremove_all.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				DefaultTableModel model = (DefaultTableModel) table.getModel();
				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}
				hashset.clear();
				filesno = 1;
				countforfile = 0;
				btn_Next.setEnabled(false);
				btn_Next.setVisible(false);
				btn_remove.setEnabled(false);
				btnremove_all.setEnabled(false);

			}

		});
		btnremove_all.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_select_folder.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				files = null;
				btnfolder = true;

				btn_Next.setVisible(true);
				btn_Next.setEnabled(true);

				btn_remove.setEnabled(true);
				btnremove_all.setEnabled(true);

				jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

				jFileChooser.setMultiSelectionEnabled(true);

				jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				jFileChooser.showOpenDialog(restore.this);

				files = jFileChooser.getSelectedFiles();

				for (int i = 0; i < files.length; i++) {
					hashset.add(files[i]);
				}

				DefaultTableModel model = (DefaultTableModel) table.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
						filesno--;
					}
				}

				Iterator<File> itr = hashset.iterator();
				while (itr.hasNext()) {
					modeli = (DefaultTableModel) table.getModel();
					File fo = itr.next();

					String filet = "";
					if (fo.isFile()) {
						filet = "File";
					} else {
						filet = "Folder";
					}
					long sizeInBytes = fo.length();
					modeli.addRow(new Object[] { "<html><b>" + filesno, "<html><b>" + fo.getName(),
							"<html><b>" + fo.getAbsolutePath(), "<html><b>" + filet,
							"<html><b>" + bytes2String(sizeInBytes) });
					filesno++;
					countforfile++;
				}
				if (countforfile == 0) {

					btn_Next.setEnabled(false);
					btn_Next.setVisible(false);
					btn_remove.setEnabled(false);
					btnremove_all.setEnabled(false);
				}

			}

		});

		JPanel panel_2 = new JPanel();
		panel_2.setBackground(Color.WHITE);
		Cardlayout.add(panel_2, "panel_2");
		panel_2.setLayout(null);

		JButton btnViewer = new JButton("");
		btnViewer.setToolTipText("Click here to See Preview.");
		btnViewer.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
			}
		});

		btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
		btnViewer.setRolloverEnabled(false);
		btnViewer.setRequestFocusEnabled(false);
		btnViewer.setOpaque(false);
		btnViewer.setFocusable(false);
		btnViewer.setFocusTraversalKeysEnabled(false);
		btnViewer.setFocusPainted(false);
		btnViewer.setDefaultCapable(false);
		btnViewer.setContentAreaFilled(false);
		btnViewer.setBorderPainted(false);
		btnViewer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (contatcheck) {

					CardLayout card1 = (CardLayout) innercardlayout.getLayout();
					card1.show(innercardlayout, "panel_Contact");

				} else if (calendarcheck) {
					CardLayout card1 = (CardLayout) innercardlayout.getLayout();
					card1.show(innercardlayout, "panel_Callendar");
				} else {
					CardLayout card = (CardLayout) innercardlayout.getLayout();

					card.show(innercardlayout, "viewer");
				}

			}
		});

		label_10 = new JLabel("");
		label_10.setRequestFocusEnabled(false);
		label_10.setFocusable(false);
		label_10.setFocusTraversalKeysEnabled(false);
		label_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));
		label_10.setBounds(10, 569, 791, 38);
		label_10.setVisible(false);
		panel_2.add(label_10);
		btnViewer.setFont(new Font("Tahoma", Font.BOLD, 13));
		btnViewer.setBounds(623, 70, 123, 38);
		panel_2.add(btnViewer);

		JButton btnAttachment = new JButton("");
		btnAttachment.setToolTipText("Click here to See Attachment.");
		btnAttachment.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
			}
		});

		btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
		btnAttachment.setDefaultCapable(false);
		btnAttachment.setFocusTraversalKeysEnabled(false);
		btnAttachment.setRolloverEnabled(false);
		btnAttachment.setRequestFocusEnabled(false);
		btnAttachment.setOpaque(false);
		btnAttachment.setFocusable(false);
		btnAttachment.setFocusPainted(false);
		btnAttachment.setContentAreaFilled(false);
		btnAttachment.setBorderPainted(false);
		btnAttachment.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				CardLayout card = (CardLayout) innercardlayout.getLayout();
				card.show(innercardlayout, "attachmenttable");

			}
		});
		btnAttachment.setFont(new Font("Tahoma", Font.BOLD, 13));
		btnAttachment.setBounds(739, 70, 116, 38);
		panel_2.add(btnAttachment);

		lblLoadingPleaseWait = new JLabel("Please Wait...");
		lblLoadingPleaseWait.setFont(new Font("Tahoma", Font.BOLD | Font.ITALIC, 9));
		lblLoadingPleaseWait.setVisible(false);

		btn_cancel = new JButton("");
		btn_cancel.setToolTipText("Click here to Stop the Process.");
		btn_cancel.setVisible(false);
		btn_cancel.setEnabled(false);
		btn_cancel.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btn_cancel.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		btn_cancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(restore.this, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					Stoppreview = true;
					lblLoadingPleaseWait.setVisible(false);
					label_10.setVisible(false);
					btnAttachment.setEnabled(true);
					table_fileinformation.setEnabled(true);
					btnViewer.setEnabled(true);
					btn_next_pane2.setEnabled(true);
					btn_previous_p2.setEnabled(true);
				}

			}
		});
		btn_cancel.setBounds(870, 78, 89, 23);
		panel_2.add(btn_cancel);
		lblLoadingPleaseWait.setBounds(996, 70, 79, 23);
		panel_2.add(lblLoadingPleaseWait);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setBounds(0, 0, 295, 549);
		panel_2.add(scrollPane_1);

		JLabel lblNew_setemail = new JLabel("");
		lblNew_setemail.setBounds(683, 4, 182, 20);
		panel_2.add(lblNew_setemail);

		JLabel lblNew_setsubject = new JLabel("");
		lblNew_setsubject.setBounds(688, 39, 377, 16);
		panel_2.add(lblNew_setsubject);

		JLabel label_date = new JLabel("");
		label_date.setBounds(935, 8, 123, 16);
		panel_2.add(label_date);

		tree = new CheckboxTree();
		scrollPane_1.setViewportView(tree);
		tree.addMouseListener(new MouseAdapter() {

			public void mouseClicked(MouseEvent arg0) {
				th = new Thread(new Runnable() {

					public void run() {

						try {

							CardLayout card = (CardLayout) innercardlayout.getLayout();

							card.show(innercardlayout, "viewer");
							editorPane.setText("");
							lblLoadingPleaseWait.setVisible(true);
							label_10.setVisible(true);
							listmail.clear();
							Stoppreview = false;
							listmapi.clear();
							lblTotalMessageCount.setText("Total Message Count : ");
							btn_cancel.setVisible(true);
							btn_cancel.setEnabled(true);
							listExchangemesingo.clear();
							listImapmesinfo.clear();
							listPSTOSTgemesingo.clear();
							btnAttachment.setEnabled(false);
							btnViewer.setEnabled(false);
							btn_next_pane2.setEnabled(false);
							btn_previous_p2.setEnabled(false);

							if (arg0.getClickCount() == 2) {
								TreePath tp = tree.getSelectionPath();

								DefaultMutableTreeNode node = (DefaultMutableTreeNode) tp.getLastPathComponent();

								foldername = node.getUserObject().toString();

								DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

								while (model.getRowCount() > 0) {

									for (int i = 0; i < model.getRowCount(); ++i) {

										model.removeRow(i);
									}
								}

								DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

								while (model1.getRowCount() > 0) {

									for (int i = 0; i < model1.getRowCount(); ++i) {

										model1.removeRow(i);
									}
								}
								editorPane.setText("");

								lblNew_setemail.setText("");

								lblNew_setsubject.setText("");

								label_date.setText("");

								if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")
										|| fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
									foldername = "";

									TreeNode[] folder = node.getPath();

									for (int i = 0; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										s = s.replace("<html><b>", "");
										if (i == 2) {
											path2 = s;
										}
										if (i == 3) {
											foldername = s;

										} else if (i > 3) {
											foldername = foldername + File.separator + s;
										}

									}
									if (!foldername.equalsIgnoreCase(root.toString())) {

										try {
											pst = PersonalStorage.fromFile(path2);
											fileInhformation_Ost_Pst();
										} catch (Exception e) {
											JOptionPane.showMessageDialog(mf,
													"File is Currupted  Please Choose another file  " + filepath,
													messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
													new ImageIcon(Main_Frame.class.getResource("/information.png")));

										}

									}

								}
							}
							lblLoadingPleaseWait.setVisible(false);
							label_10.setVisible(false);
							btn_cancel.setVisible(false);
							btn_cancel.setEnabled(false);
							btnAttachment.setEnabled(true);
							table_fileinformation.setEnabled(true);
							btnViewer.setEnabled(true);
							btn_next_pane2.setEnabled(true);
							btn_previous_p2.setEnabled(true);

						} catch (Exception e) {

						} finally {

						}

					}

				});
				th.start();

			}
		});

		tree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("root") {

			private static final long serialVersionUID = 1L;
			{

			}
		}));

		btn_next_pane2 = new JButton("");
		btn_next_pane2.setToolTipText("Click here to Go Forward.");
		btn_next_pane2.setDefaultCapable(false);
		btn_next_pane2.setContentAreaFilled(false);
		btn_next_pane2.setBorderPainted(false);
		btn_next_pane2.setRequestFocusEnabled(false);
		btn_next_pane2.setOpaque(false);
		btn_next_pane2.setRolloverEnabled(false);
		btn_next_pane2.setFocusable(false);
		btn_next_pane2.setFocusTraversalKeysEnabled(false);
		btn_next_pane2.setFocusPainted(false);
		btn_next_pane2.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
			}
		});

		btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));

		btn_next_pane2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				TreePath[] tp = tree.getCheckingPaths();

				TreePath[] checktp11 = tree.getCheckingPaths();
				if (checktp11.length == 0) {
					JOptionPane.showMessageDialog(mf, "Select File From the Tree", messageboxtitle,
							JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				} else {

					hm = new HashMap<String, List<String>>();
					textField_importcsvpath.setText("");
					pstfolderlist = null;
					destination_path = "";
					for (int i = 0; i < tp.length; i++) {

						pstfolderlist2 = new ArrayList<String>();
						String[] str = (tp[i].toString().replace("<html><b>", "")).split(",");
						String sfile = "";
						StringBuilder strbr = new StringBuilder();
						for (int j = 3; j < str.length; j++) {

							if (j != (str.length - 1)) {
								strbr.append(str[j].trim());
								if (!pstfolderlist2.contains(strbr.toString().trim())) {
									pstfolderlist2.add(strbr.toString().trim());
								}
								strbr.append(File.separator);

							} else if (j == str.length - 1) {
								strbr.append(str[j].replace("]", "").trim());
								if (!pstfolderlist2.contains(strbr.toString().trim())) {
									pstfolderlist2.add(strbr.toString().trim());
								}
							}

						}

						if (!(fileoptionm.equalsIgnoreCase("MBOX") || fileoptionm.equalsIgnoreCase("EML Files (.eml)")
								|| fileoptionm.equalsIgnoreCase("Message Files (.msg)")
								|| fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)"))) {

							String fname;
							try {
								fname = str[2].toString();
							} catch (Exception e33) {
								continue;
							}

							if (hm.size() != 0) {
								pstfolderlist = hm.get(fname);
								if (pstfolderlist != null) {
									if (!pstfolderlist.contains(strbr.toString().trim())) {
										pstfolderlist.add(strbr.toString().trim());
										hm.put(fname, pstfolderlist);
									}

								} else {
									pstfolderlist = new ArrayList<String>();

									pstfolderlist.addAll(pstfolderlist2);
									if (!pstfolderlist.contains(strbr.toString().trim())) {
										pstfolderlist.add(strbr.toString().trim());
									}

									hm.put(fname, pstfolderlist);
								}

							} else {
								pstfolderlist = new ArrayList<String>();
								pstfolderlist.add(strbr.toString().trim());
								hm.put(fname, pstfolderlist);
							}
						} else {
							DefaultMutableTreeNode d1 = (DefaultMutableTreeNode) tp[i].getLastPathComponent();
							if (d1.isLeaf()) {
								sfile = str[2].replace("]", "").trim();
								sfile = sfile + str[3].replace("]", "").trim();

								hm.put(sfile, null);
							}
						}

					}

					panel_progress.setVisible(false);

					if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

						comboBox_fileDestination_type.removeItem("OST");
					} else if (fileoptionm.equalsIgnoreCase("EML Files (.eml)")) {

						comboBox_fileDestination_type.removeItem("EML");
					} else if (fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")) {

						comboBox_fileDestination_type.removeItem("EMLX");

					} else if (fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {

						comboBox_fileDestination_type.removeItem("MSG");

					} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
						comboBox_fileDestination_type.removeItem("MBOX");
						comboBox_fileDestination_type.removeItem("Thunderbird");
						comboBox_fileDestination_type.removeItem("Opera Mail");
					}

					panel_3_2.setVisible(false);
					panel_3_.setVisible(false);
					panel_3_1_2.setVisible(false);
					btn_converter_1.setEnabled(false);
					comboBox.setVisible(false);
					btnStop.setVisible(false);
					DefaultTableModel model = (DefaultTableModel) table_userid.getModel();
					while (model.getRowCount() > 0) {

						for (int i = 0; i < model.getRowCount(); ++i) {

							model.removeRow(i);
						}
					}
					tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");

					lbl_progressreport.setText("");

					panel_3_.setVisible(true);

					CardLayout card1 = (CardLayout) panel_3_.getLayout();
					card1.show(panel_3_, "panel_3_1_1");

					panel_progress.setVisible(true);
					filetype = "OFFICE 365";
					btn_converter_1.setEnabled(true);

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_mailboxid");

					sigin sig = new sigin(mf);
					sig.setLocationRelativeTo(mf);
					sig.setVisible(true);
					password_p3 = sigin.password_p3;
					username_p3 = sigin.username_p3;
					domain_p3 = sigin.domain_p3;
					fileoptioncheck = sigin.check;
					if (!(username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase(""))) {
						SwingWorker sw1 = new SwingWorker() {
							@Override
							protected Object doInBackground() {

								obThmailbox = new loadingThreadclassformailbox(mf);
								obThmailbox.start();

								if (fileoptionm.equalsIgnoreCase("Exchange Backup & Restore")) {

									mailboxUri = "https://" + domain_p3 + "/ews/Exchange.asmx";
								}

								try {
									conntiontooffice365_output();

								} catch (Exception e) {

								}

								usernameverification = username_p3.split("@");
								obThmailbox.ob.MessageLabel.setText("Getting Mailbox User Id Please wait...");
								if (sig.chckbxIMAdmin.isSelected()) {
									impersonationUser();
								} else {
									DefaultTableModel dtm = (DefaultTableModel) table_userid.getModel();
									dtm.addRow(new Object[] { true, "", username_p3 });
								}

								return null;
							}

							@Override
							protected void done() {

								obThmailbox.close();
							}
						};

						sw1.execute();

						sig.dispose();

						cal = Calendar.getInstance();
						calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
					}
				}

			}
		});
		btn_next_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));
		btn_next_pane2.setBounds(933, 567, 123, 38);
		panel_2.add(btn_next_pane2);

		btn_previous_p2 = new JButton("");
		btn_previous_p2.setToolTipText("Click here to Go Back.");
		btn_previous_p2.setRolloverEnabled(false);
		btn_previous_p2.setRequestFocusEnabled(false);
		btn_previous_p2.setOpaque(false);
		btn_previous_p2.setFocusable(false);
		btn_previous_p2.setFocusTraversalKeysEnabled(false);
		btn_previous_p2.setFocusPainted(false);
		btn_previous_p2.setDefaultCapable(false);
		btn_previous_p2.setContentAreaFilled(false);
		btn_previous_p2.setBorderPainted(false);
		btn_previous_p2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous_p2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous_p2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous_p2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous_p2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				lblNew_setsubject.setText("");
				lblNew_setemail.setText("");
				label_date.setText("");
				editorPane.setText("");
				lblTotalMessageCount.setText("");
				model = (DefaultTreeModel) tree.getModel();
				DefaultMutableTreeNode root = (DefaultMutableTreeNode) model.getRoot();
				root.removeAllChildren();
				model.reload();
				TreePath[] ac = new TreePath[0];
				tree.setCheckingPaths(ac);
				DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

				while (model1.getRowCount() > 0) {

					for (int i = 0; i < model1.getRowCount(); ++i) {

						model1.removeRow(i);
					}
				}

				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_1");
			}
		});
		btn_previous_p2.setFont(new Font("Tahoma", Font.BOLD, 13));
		btn_previous_p2.setBounds(811, 567, 123, 38);
		panel_2.add(btn_previous_p2);

		JScrollPane scrollPane_2 = new JScrollPane();
		scrollPane_2.setBounds(298, 0, 324, 527);
		panel_2.add(scrollPane_2);

		innercardlayout = new JPanel();
		innercardlayout.setBounds(627, 106, 448, 443);
		panel_2.add(innercardlayout);
		innercardlayout.setLayout(new CardLayout(0, 0));

		CardLayout card = (CardLayout) innercardlayout.getLayout();
		card.show(innercardlayout, "viewer");

		table_fileinformation = new JTable() {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};

		table_fileinformation.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try {
					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							if (arg0.getClickCount() == 1) {
								editorPane.setText("");
								DefaultTableModel model = (DefaultTableModel) table_1.getModel();
								while (model.getRowCount() > 0) {

									for (int i = 0; i < model.getRowCount(); ++i) {

										model.removeRow(i);
									}
								}
								contatcheck = false;
								calendarcheck = false;
								if (fileoptionm.equalsIgnoreCase("MBOX")) {
									MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

									try {
										lblNew_setemail.setText(message.getFrom().toString());

										lblNew_setsubject.setText(message.getSubject());

										label_date.setText(message.getDate().toString());

										HTMLEditorKit kit = new HTMLEditorKit();
										editorPane.setEditorKit(kit);
										FileOutputStream os = new FileOutputStream(
												temppathm + File.separator + "previewHtml.html");
										message.save(os, EmlSaveOptions.getDefaultHtml());
										os.close();
										URL url = new URL("file:///" + temppathm + File.separator + "previewHtml.html");
										editorPane.setPage(url);

										for (int j = 0; j < message.getAttachments().size(); j++) {
											Attachment att = message.getAttachments().get_Item(j);

											String attFileName = att.getName();
											ImageIcon icon = null;

											if (attFileName.endsWith(".pdf")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
											} else if (attFileName.endsWith(".txt")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
											} else if (attFileName.endsWith(".docx")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
											} else if (attFileName.endsWith(".zip")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
											} else {
												icon = new ImageIcon(
														Main_Frame.class.getResource("/attachment-icon.png"));
											}
											JLabel imagelabl = new JLabel();
											imagelabl.setIcon(icon);
											DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
											modeli.addRow(new Object[] { "<html><b>" + (j + 1),
													"<html><b>" + attFileName, imagelabl });

										}

									} catch (Exception e) {

									}

								} else if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

									MapiMessage message = listmapi.get(table_fileinformation.getSelectedRow());

									try {
										lblNew_setemail.setText(message.getSenderEmailAddress());
									} catch (Exception a) {
										lblNew_setemail.setText("");
									}
									try {
										lblNew_setsubject.setText(message.getSubject());
									} catch (Exception a) {
										lblNew_setsubject.setText("");
									}
									try {
										label_date.setText(message.getDeliveryTime().toString());
									} catch (Exception a) {
										label_date.setText("");
									}

									if (message.getMessageClass().equals("IPM.Contact")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Contact");

										MapiContact con = (MapiContact) message.toMapiMessageItem();
										try {
											String[] compa = con.getCompanies();

											label_contactcompany.setText(compa[0]);
										} catch (Exception e) {
											label_contactcompany.setText("");
										}
										try {
											String gn = "";
											String mn = "";
											String sn = "";
											try {
												gn = con.getNameInfo().getGivenName();
											} catch (Exception e) {

											}
											try {
												mn = con.getNameInfo().getMiddleName();
											} catch (Exception e) {

											}
											try {
												sn = con.getNameInfo().getSurname();
											} catch (Exception e) {

											}

											String fn = gn + " " + mn + " " + sn;
											label_contactfullname.setText(fn);
										} catch (Exception e) {
											label_contactfullname.setText("");
										}
										try {
											label_contactemail.setText(
													con.getElectronicAddresses().getEmail1().getEmailAddress());
										} catch (Exception e) {
											label_contactemail.setText("");
										}
										try {
											label_contactphonenumber
													.setText(con.getTelephones().getMobileTelephoneNumber());
										} catch (Exception e) {
											label_contactphonenumber.setText("");
										}
										try {
											textArea_contact.setText(con.getPersonalInfo().getNotes());
										} catch (Exception e) {
											textArea_contact.setText("");
										}

										contatcheck = true;

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Callendar");
										System.out.println("check");
										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

										calendarcheck = true;
										try {
											label_Calendarsubject.setText(cal.getSubject());
										} catch (Exception e) {
											label_Calendarsubject.setText("");
										}
										try {
											label_calendarstartdate.setText(cal.getStartDate().toString());
										} catch (Exception e) {
											label_calendarstartdate.setText("");
										}
										try {
											label_Calendarenddate.setText(cal.getEndDate().toString());
										} catch (Exception e) {
											label_Calendarenddate.setText("");
										}

									} else {

										try {

											HTMLEditorKit kit = new HTMLEditorKit();
											editorPane.setEditorKit(kit);
											FileOutputStream os = new FileOutputStream(
													temppathm + File.separator + "previewHtml.html");
											message.save(os, EmlSaveOptions.getDefaultHtml());
											os.close();
											URL url = new URL(
													"file:///" + temppathm + File.separator + "previewHtml.html");
											editorPane.setPage(url);

										} catch (Error e) {
											mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
										} catch (Exception e) {
											mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
											editorPane.setContentType("text/html");
											editorPane.setText("<html>Page not found.</html>");
										}

									}

									int k = 1;
									for (int j = 0; j < message.getAttachments().size(); j++) {
										MapiAttachment att = message.getAttachments().get_Item(j);

										String attFileName = att.getLongFileName();
										ImageIcon icon = null;

										if (attFileName.endsWith(".pdf")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
										} else if (attFileName.endsWith(".txt")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
										} else if (attFileName.endsWith(".docx")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
										} else if (attFileName.endsWith(".zip")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
										} else {
											icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
										}
										JLabel imagelabl = new JLabel();
										imagelabl.setIcon(icon);

										DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
										modeli.addRow(
												new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
										k++;

									}

								} else if (fileoptionm.equalsIgnoreCase("OLM files (.olm)")) {

									MapiMessage message = listmapi.get(table_fileinformation.getSelectedRow());
									try {
										lblNew_setemail.setText(message.getSenderEmailAddress());
									} catch (Exception a) {
										lblNew_setemail.setText("");
									}
									try {
										lblNew_setsubject.setText(message.getSubject());
									} catch (Exception a) {
										lblNew_setsubject.setText("");
									}
									try {
										label_date.setText(message.getDeliveryTime().toString());
									} catch (Exception a) {
										label_date.setText("");
									}

									try {
										lblNew_setemail.setText(message.getSenderEmailAddress());
									} catch (Exception a) {
										lblNew_setemail.setText("");
									}
									try {
										lblNew_setsubject.setText(message.getSubject());
									} catch (Exception a) {
										lblNew_setsubject.setText("");
									}
									try {
										label_date.setText(message.getDeliveryTime().toString());
									} catch (Exception a) {
										label_date.setText("");
									}

									if (message.getMessageClass().equals("IPM.Contact")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Contact");

										MapiContact con = (MapiContact) message.toMapiMessageItem();
										try {
											String[] compa = con.getCompanies();

											label_contactcompany.setText(compa[0]);
										} catch (Exception e) {
											label_contactcompany.setText("");
										}
										try {
											String gn = "";
											String mn = "";
											String sn = "";
											try {
												gn = con.getNameInfo().getGivenName();
											} catch (Exception e) {

											}
											try {
												mn = con.getNameInfo().getMiddleName();
											} catch (Exception e) {

											}
											try {
												sn = con.getNameInfo().getSurname();
											} catch (Exception e) {

											}

											String fn = gn + " " + mn + " " + sn;
											label_contactfullname.setText(fn);
										} catch (Exception e) {
											label_contactfullname.setText("");
										}
										try {
											label_contactemail.setText(
													con.getElectronicAddresses().getEmail1().getEmailAddress());
										} catch (Exception e) {
											label_contactemail.setText("");
										}
										try {
											label_contactphonenumber
													.setText(con.getTelephones().getMobileTelephoneNumber());
										} catch (Exception e) {
											label_contactphonenumber.setText("");
										}
										try {
											textArea_contact.setText(con.getPersonalInfo().getNotes());
										} catch (Exception e) {
											textArea_contact.setText("");
										}

										contatcheck = true;

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {
										CardLayout card1 = (CardLayout) innercardlayout.getLayout();
										card1.show(innercardlayout, "panel_Callendar");
										System.out.println("check");
										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

										calendarcheck = true;
										try {
											label_Calendarsubject.setText(cal.getSubject());
										} catch (Exception e) {
											label_Calendarsubject.setText("");
										}
										try {
											label_calendarstartdate.setText(cal.getStartDate().toString());
										} catch (Exception e) {
											label_calendarstartdate.setText("");
										}
										try {
											label_Calendarenddate.setText(cal.getEndDate().toString());
										} catch (Exception e) {
											label_Calendarenddate.setText("");
										}

									} else {

										try {

											HTMLEditorKit kit = new HTMLEditorKit();
											editorPane.setEditorKit(kit);
											FileOutputStream os = new FileOutputStream(
													temppathm + File.separator + "previewHtml.html");
											message.save(os, EmlSaveOptions.getDefaultHtml());
											os.close();
											URL url = new URL(
													"file:///" + temppathm + File.separator + "previewHtml.html");
											editorPane.setPage(url);

										} catch (Error e) {
											mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
										} catch (Exception e) {
											mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
											editorPane.setContentType("text/html");
											editorPane.setText("<html>Page not found.</html>");
										}

									}

									int k = 1;
									for (int j = 0; j < message.getAttachments().size(); j++) {
										MapiAttachment att = message.getAttachments().get_Item(j);

										String attFileName = att.getLongFileName();
										ImageIcon icon = null;

										if (attFileName.endsWith(".pdf")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
										} else if (attFileName.endsWith(".txt")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
										} else if (attFileName.endsWith(".docx")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
										} else if (attFileName.endsWith(".zip")) {
											icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
										} else {
											icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
										}
										JLabel imagelabl = new JLabel();
										imagelabl.setIcon(icon);

										DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
										modeli.addRow(
												new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
										k++;

									}
								} else if (fileoptionm.equalsIgnoreCase("Nodes Storage (.nsf)")) {

								} else {
									MailMessage message = MailMessage.load(path2);

									try {
										try {
											lblNew_setemail.setText(message.getFrom().toString());
										} catch (Exception e) {

										}
										try {
											lblNew_setsubject.setText(message.getSubject());
										} catch (Exception e) {

										}
										try {
											label_date.setText(message.getDate().toString());
										} catch (Exception e) {

										}
										HTMLEditorKit kit = new HTMLEditorKit();
										editorPane.setEditorKit(kit);
										FileOutputStream os = new FileOutputStream(
												temppathm + File.separator + "previewHtml.html");
										message.save(os, EmlSaveOptions.getDefaultHtml());
										os.close();
										URL url = new URL("file:///" + temppathm + File.separator + "previewHtml.html");
										editorPane.setPage(url);

										for (int j = 0; j < message.getAttachments().size(); j++) {
											Attachment att = message.getAttachments().get_Item(j);

											String attFileName = att.getName();
											ImageIcon icon = null;

											if (attFileName.endsWith(".pdf")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
											} else if (attFileName.endsWith(".txt")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
											} else if (attFileName.endsWith(".docx")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
											} else if (attFileName.endsWith(".zip")) {
												icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
											} else {
												icon = new ImageIcon(
														Main_Frame.class.getResource("/attachment-icon.png"));
											}
											JLabel imagelabl = new JLabel();
											imagelabl.setIcon(icon);

											DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
											modeli.addRow(new Object[] { "<html><b>" + (j + 1),
													"<html><b>" + attFileName, imagelabl });
											// System.out.println(attFileName);
										}
									} catch (Exception e) {

										// e.printStackTrace();
									}
								}

							}
						}
					});

				} catch (Exception e) {

				}
			}
		});
		table_fileinformation.setModel(
				new DefaultTableModel(new Object[][] {}, new String[] { "From", "Subject", "Date", "Attachment" }));
		table_fileinformation.getColumn("Attachment").setCellRenderer(new Renderer());
		table_fileinformation.getTableHeader().setReorderingAllowed(false);
		scrollPane_2.setViewportView(table_fileinformation);

		JLabel lblNewLabel_2 = new JLabel("From");
		lblNewLabel_2.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_2.setBounds(628, 5, 45, 23);
		panel_2.add(lblNewLabel_2);

		JLabel lblNewLabel_3 = new JLabel("Subject");
		lblNewLabel_3.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_3.setBounds(627, 39, 63, 16);
		panel_2.add(lblNewLabel_3);

		JLabel lblDate = new JLabel("Date");
		lblDate.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblDate.setBounds(882, 13, 45, 16);
		panel_2.add(lblDate);

		JPanel viewer = new JPanel();
		innercardlayout.add(viewer, "viewer");
		viewer.setLayout(null);

		scrollPane_3 = new JScrollPane();
		scrollPane_3.setBounds(0, 0, 448, 443);
		viewer.add(scrollPane_3);

		editorPane = new JEditorPane();
		scrollPane_3.setViewportView(editorPane);
		editorPane.setEditable(false);

		panel_Callendar = new JPanel();
		panel_Callendar.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Callendar.setBackground(Color.WHITE);
		innercardlayout.add(panel_Callendar, "panel_Callendar");
		panel_Callendar.setLayout(null);

		label_calendarstartdate = new JLabel("");
		label_calendarstartdate.setBounds(78, 151, 351, 17);
		panel_Callendar.add(label_calendarstartdate);

		lblEndDate = new JLabel("End Date");
		lblEndDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblEndDate.setBounds(10, 199, 65, 17);
		panel_Callendar.add(lblEndDate);

		JLabel lblStartDate = new JLabel("Start Date");
		lblStartDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblStartDate.setBounds(10, 148, 65, 22);
		panel_Callendar.add(lblStartDate);

		JLabel lblSubject = new JLabel("Subject");
		lblSubject.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblSubject.setBounds(10, 92, 52, 28);
		panel_Callendar.add(lblSubject);

		label_Calendarsubject = new JLabel("");
		label_Calendarsubject.setBounds(85, 92, 344, 28);
		panel_Callendar.add(label_Calendarsubject);

		label_Calendaricon = new JLabel("");
		label_Calendaricon.setIcon(new ImageIcon(Main_Frame.class.getResource("/calender.png")));
		label_Calendaricon.setBounds(10, 11, 116, 77);
		panel_Callendar.add(label_Calendaricon);

		label_Calendarenddate = new JLabel("");
		label_Calendarenddate.setBounds(78, 199, 351, 22);
		panel_Callendar.add(label_Calendarenddate);

		panel_Contact = new JPanel();
		panel_Contact.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Contact.setBackground(Color.WHITE);
		innercardlayout.add(panel_Contact, "panel_Contact");
		panel_Contact.setLayout(null);

		lblFullName = new JLabel("Full Name ");
		lblFullName.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblFullName.setBounds(10, 67, 64, 14);
		panel_Contact.add(lblFullName);

		label_contactfullname = new JLabel("");
		label_contactfullname.setBounds(83, 67, 215, 14);
		panel_Contact.add(label_contactfullname);

		lblEMailAd = new JLabel("E Mail :");
		lblEMailAd.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblEMailAd.setBounds(10, 92, 64, 21);
		panel_Contact.add(lblEMailAd);

		lblCompany = new JLabel("Company");
		lblCompany.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblCompany.setBounds(10, 124, 64, 14);
		panel_Contact.add(lblCompany);

		label_contactemail = new JLabel("");
		label_contactemail.setBounds(83, 95, 215, 18);
		panel_Contact.add(label_contactemail);

		label_contactcompany = new JLabel("");
		label_contactcompany.setBounds(83, 124, 215, 21);
		panel_Contact.add(label_contactcompany);

		lblPhoneNo = new JLabel("Phone No");
		lblPhoneNo.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblPhoneNo.setBounds(10, 149, 56, 21);
		panel_Contact.add(lblPhoneNo);

		label_contactphonenumber = new JLabel("");
		label_contactphonenumber.setBounds(80, 152, 218, 18);
		panel_Contact.add(label_contactphonenumber);

		textArea_contact = new JTextArea();
		textArea_contact.setBounds(10, 198, 419, 212);
		panel_Contact.add(textArea_contact);

		label_contacticon = new JLabel("");
		label_contacticon.setIcon(new ImageIcon(Main_Frame.class.getResource("/User-Chat-icon.png")));
		label_contacticon.setBounds(10, 11, 64, 51);
		panel_Contact.add(label_contacticon);

		lblNotes = new JLabel("Notes");
		lblNotes.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNotes.setBounds(10, 172, 48, 21);
		panel_Contact.add(lblNotes);

		attachmenttable = new JPanel();
		innercardlayout.add(attachmenttable, "attachmenttable");
		attachmenttable.setLayout(null);

		scrollPane_4 = new JScrollPane();
		scrollPane_4.setBounds(0, 0, 448, 443);
		attachmenttable.add(scrollPane_4);
		Object[][] data12 = {};

		String[] cols12 = { "S.No", "File Name", "File Type" };

		DefaultTableModel tablemodel2 = new DefaultTableModel(data12, cols12) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {
				// all cells false

				return false;

			}

		};
		table_1 = new JTable(tablemodel2);
		scrollPane_4.setViewportView(table_1);

		JLabel label_6 = new JLabel("");
		label_6.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_6.setBounds(0, 552, 1129, 66);
		panel_2.add(label_6);

		lblTotalMessageCount = new JLabel("Total Message Count :");
		lblTotalMessageCount.setBounds(298, 529, 324, 23);
		panel_2.add(lblTotalMessageCount);
		table_1.getColumn("File Type").setCellRenderer(new Renderer());

		panel_mailboxid = new JPanel();
		panel_mailboxid.setBackground(Color.WHITE);
		Cardlayout.add(panel_mailboxid, "panel_mailboxid");
		panel_mailboxid.setLayout(null);

		JLabel lblSampleCsv = new JLabel("Sample CSV");
		lblSampleCsv.setForeground(Color.RED);
		Cursor cursor = new Cursor(Cursor.HAND_CURSOR);
		lblSampleCsv.setCursor(cursor);
		lblSampleCsv.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblSampleCsv.setBounds(519, 65, 85, 23);
		lblSampleCsv.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				File file = new File(System.getProperty("user.home") + File.separator + "Desktop" + File.separator
						+ messageboxtitle + "_sample.csv");

				try {

					if (!file.isFile()) {
						FileWriter outputfile = new FileWriter(file);

						CSVWriter writer = new CSVWriter(outputfile);

						String[] data1 = { "name@comapanyname.onmicrosoft.com", "name@schoolname.onmicrosoft.com",
								"xyz@xyz.onmicrosoft.com", "xxx@xxx.onmicrosoft.com", "abc@abcd.onmicrosoft.com" };

						for (int i = 0; i < data1.length; i++) {
							String[] data2 = { data1[i] };
							writer.writeNext(data2);
						}

						writer.close();
						file.setReadOnly();
					}

					Desktop desktop = Desktop.getDesktop();
					desktop.open(file);

				} catch (Exception e1) {
				}
			}
		});

		panel_mailboxid.add(lblSampleCsv);

		JScrollPane scrollPane_5 = new JScrollPane();
		scrollPane_5.setBounds(0, 99, 1075, 436);
		panel_mailboxid.add(scrollPane_5);

		Object[][] data1 = {};

		String[] cols1 = { "", "", "Mailbox User Id" };

		DefaultTableModel tablemodel = new DefaultTableModel(data1, cols1) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {
				// all cells false
				if (column == 0) {
					return true;
				} else {
					return false;
				}
			}

			public Class<?> getColumnClass(int column) {
				if (column == 0) {
					return Boolean.class;
				} else {
					return String.class;
				}
			}
		};

		table_userid = new JTable(tablemodel);

		table_userid.setFont(new Font("Calibri", Font.PLAIN, 14));
		// table_userid.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		table_userid.getTableHeader().setReorderingAllowed(false);
		table_userid.getTableHeader().setFont(new Font("Calibri", Font.BOLD, 14));

		DefaultTableCellRenderer tablerenderer = (DefaultTableCellRenderer) table_userid.getTableHeader()
				.getDefaultRenderer();
		tablerenderer.setHorizontalAlignment(0);
		TableColumnModel csvcolumnModel = table_userid.getColumnModel();

		csvcolumnModel.getColumn(1).setCellRenderer(new ImageRenderer());
		csvcolumnModel.getColumn(1).setPreferredWidth(20);
		csvcolumnModel.getColumn(1).setMaxWidth(20);
		csvcolumnModel.getColumn(1).setHeaderRenderer(new ImageRenderer());
		csvcolumnModel.getColumn(0).setHeaderRenderer(new HeaderRenderer2(table_userid.getTableHeader(), 0));
		csvcolumnModel.getColumn(0).setPreferredWidth(30);
		csvcolumnModel.getColumn(0).setMaxWidth(30);
		csvcolumnModel.getColumn(2).setPreferredWidth(620);

		tablerenderer.setHorizontalAlignment(0);

		scrollPane_5.setViewportView(table_userid);

		textField_importcsvpath = new JTextField();
		textField_importcsvpath.setEditable(false);
	
		textField_importcsvpath.setBounds(10, 21, 597, 33);
		panel_mailboxid.add(textField_importcsvpath);
		textField_importcsvpath.setColumns(10);

		JButton btnImportCsv = new JButton("");
		btnImportCsv.setContentAreaFilled(false);
		btnImportCsv.setBorderPainted(false);
		btnImportCsv.setRolloverEnabled(false);
		btnImportCsv.setRequestFocusEnabled(false);
		btnImportCsv.setOpaque(false);
		btnImportCsv.setFocusable(false);
		btnImportCsv.setFocusTraversalKeysEnabled(false);
		btnImportCsv.setFocusPainted(false);
		btnImportCsv.setDefaultCapable(false);
		btnImportCsv.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnImportCsv.setIcon(new ImageIcon(backup.class.getResource("/import-csv-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnImportCsv.setIcon(new ImageIcon(backup.class.getResource("/import-csv-btn.png")));
			}
		});

		btnImportCsv.setIcon(new ImageIcon(backup.class.getResource("/import-csv-btn.png")));
		btnImportCsv.setToolTipText("Click here to Import CSV File containing All the Mailbox User Id ");
		btnImportCsv.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
				chooser.setMultiSelectionEnabled(true);
				chooser.setAcceptAllFileFilterUsed(false);
				chooser.setFileFilter(new FileNameExtensionFilter("CSV File", "csv"));
				int ret = chooser.showOpenDialog(restore.this);

				Executors.newSingleThreadExecutor().execute(new Runnable() {
					public void run() {

						if (ret == JFileChooser.APPROVE_OPTION) {
							File file = new File(chooser.getSelectedFile() + "");
							if (file.exists()) {
								textField_importcsvpath.setText(file.getAbsolutePath());
								readDataLineByLine(file.getAbsolutePath());
							} else {
								JOptionPane.showMessageDialog(restore.this, "  File doesn't exists!!", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(backup.class.getResource("/information.png")));
								return;
							}
						}
					}
				});
			}
		});
		btnImportCsv.setBounds(630, 21, 142, 38);
		panel_mailboxid.add(btnImportCsv);

		JButton selectAll = new JButton("");
		selectAll.setRequestFocusEnabled(false);
		selectAll.setRolloverEnabled(false);
		selectAll.setOpaque(false);
		selectAll.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				for (int i = 0; i < table_userid.getRowCount(); i++) {
					table_userid.setValueAt(true, i, 0);

				}
			}
		});
		selectAll.setFocusable(false);
		selectAll.setToolTipText("Click here to Select all the Mailbox User Id ");
		selectAll.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				selectAll.setIcon(new ImageIcon(backup.class.getResource("/select-all-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				selectAll.setIcon(new ImageIcon(backup.class.getResource("/select-all-btn.png")));
			}
		});

		selectAll.setIcon(new ImageIcon(backup.class.getResource("/select-all-btn.png")));

		selectAll.setFocusTraversalKeysEnabled(false);
		selectAll.setFocusPainted(false);
		selectAll.setDefaultCapable(false);
		selectAll.setContentAreaFilled(false);
		selectAll.setBorderPainted(false);
		selectAll.setBounds(777, 21, 142, 38);
		panel_mailboxid.add(selectAll);

		JButton unselectall = new JButton("");
		unselectall.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				unselectall.setIcon(new ImageIcon(backup.class.getResource("/unselect-all-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				unselectall.setIcon(new ImageIcon(backup.class.getResource("/unselect-all-btn.png")));
			}
		});

		unselectall.setIcon(new ImageIcon(backup.class.getResource("/unselect-all-btn.png")));

		unselectall.setRolloverEnabled(false);
		unselectall.setToolTipText("Click here to De-Select all the Mailbox User Id ");
		unselectall.setRequestFocusEnabled(false);
		unselectall.setOpaque(false);
		unselectall.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				for (int i = 0; i < table_userid.getRowCount(); i++) {
					table_userid.setValueAt(false, i, 0);

				}
			}
		});
		unselectall.setFocusTraversalKeysEnabled(false);
		unselectall.setFocusable(false);
		unselectall.setFocusPainted(false);
		unselectall.setDefaultCapable(false);
		unselectall.setContentAreaFilled(false);
		unselectall.setBorderPainted(false);
		unselectall.setBounds(923, 21, 142, 38);
		panel_mailboxid.add(unselectall);

		JButton btnNext = new JButton("");
		btnNext.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				userList.clear();
				DefaultTableModel model = (DefaultTableModel) table_2.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				DefaultTableModel model1 = (DefaultTableModel) table_fileselect.getModel();

				while (model1.getRowCount() > 0) {

					for (int i = 0; i < model1.getRowCount(); ++i) {

						model1.removeRow(i);
					}
				}

				uservaluescomboBox.removeAllItems();
				for (int i = 0; i < table_userid.getRowCount(); i++) {
					String useridchecked = null;
					try {

						useridchecked = table_userid.getValueAt(i, 0).toString();

						if (useridchecked.equalsIgnoreCase("True")) {

							String userid = table_userid.getValueAt(i, 2).toString();
							userList.add(userid);
							uservaluescomboBox.addItem(userid);

						}
					} catch (Exception e1) {

					}

				}
				if (!userList.isEmpty()) {

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_fileselect");

					DefaultTableModel yourModel_1 = (DefaultTableModel) table_fileselect.getModel();

					for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
						for (int j = 0; j < filesfin.length; j++) {
							File filenamed = new File(filesfin[j].replace("<html><b>", ""));

							if (filenamed.isFile()) {
								if (entry.getKey().trim().equalsIgnoreCase(filenamed.getAbsolutePath())) {
									yourModel_1
											.addRow(new Object[] { new File(entry.getKey()).getName(), username_p3 });
								}
							} else {
								try {
									readafile(new File(filesfin[j].replace("<html><b>", "")), entry, yourModel_1);
								} catch (Exception e) {

									e.printStackTrace();
								}
							}

						}

					}
				} else {
					JOptionPane.showMessageDialog(mf, "Please Select the Mailobox before continuing", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));
				}
			}
		});
		btnNext.setRolloverEnabled(false);
		btnNext.setRequestFocusEnabled(false);
		btnNext.setOpaque(false);
		btnNext.setFocusable(false);

		btnNext.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnNext.setIcon(new ImageIcon(backup.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnNext.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));
			}
		});

		btnNext.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));

		btnNext.setFocusTraversalKeysEnabled(false);
		btnNext.setFocusPainted(false);
		btnNext.setDefaultCapable(false);
		btnNext.setToolTipText("Click here to Go To the tree panel ");
		btnNext.setContentAreaFilled(false);
		btnNext.setBorderPainted(false);
		btnNext.setBounds(931, 567, 142, 38);
		panel_mailboxid.add(btnNext);

		JButton btn_previous1 = new JButton("");
		btn_previous1.setToolTipText("Click here to Go Back.");
		btn_previous1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DefaultTableModel model = (DefaultTableModel) table_userid.getModel();
				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_2");
			}
		});
		btn_previous1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous1.setIcon(new ImageIcon(backup.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous1.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous1.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
		btn_previous1.setRolloverEnabled(false);
		btn_previous1.setRequestFocusEnabled(false);
		btn_previous1.setOpaque(false);
		btn_previous1.setFocusable(false);
		btn_previous1.setFocusTraversalKeysEnabled(false);
		btn_previous1.setFocusPainted(false);
		btn_previous1.setDefaultCapable(false);
		btn_previous1.setContentAreaFilled(false);
		btn_previous1.setBorderPainted(false);
		btn_previous1.setBounds(811, 567, 142, 38);
		panel_mailboxid.add(btn_previous1);

		lblNewLabel_5 = new JLabel("");
		lblNewLabel_5.setBounds(0, 552, 1075, 68);
		lblNewLabel_5.setIcon(new ImageIcon(backup.class.getResource("/bottom.png")));
		panel_mailboxid.add(lblNewLabel_5);

		panel_fileselect = new JPanel();
		panel_fileselect.setBackground(Color.WHITE);
		Cardlayout.add(panel_fileselect, "panel_fileselect");
		panel_fileselect.setLayout(null);

		JButton button_1 = new JButton("");
		button_1.setToolTipText("Click here to Go Back.");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DefaultTableModel model = (DefaultTableModel) table_fileselect.getModel();
				for (int i = 0; i < table_fileselect.getRowCount(); i++) {
					model.removeRow(i);

				}
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_mailboxid");
			}
		});
		button_1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button_1.setIcon(new ImageIcon(backup.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				button_1.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
			}
		});

		button_1.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
		button_1.setRolloverEnabled(false);
		button_1.setRequestFocusEnabled(false);
		button_1.setOpaque(false);
		button_1.setFocusable(false);
		button_1.setFocusTraversalKeysEnabled(false);
		button_1.setFocusPainted(false);
		button_1.setDefaultCapable(false);
		button_1.setContentAreaFilled(false);
		button_1.setBorderPainted(false);
		button_1.setBounds(800, 567, 142, 38);

		panel_fileselect.add(button_1);

		JButton button = new JButton("");
		button.setToolTipText("Click here to Go Forward.");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				DefaultTableModel model = (DefaultTableModel) table_2.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				SwingWorker sw1 = new SwingWorker() {
					@Override
					protected Object doInBackground() {
						obTh = new LoadingThreadclass(mf);
						obTh.start();

						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_mailfolder");

						DefaultTableModel dtm = (DefaultTableModel) table_2.getModel();

						TableColumn testColumn = csvcolumnModel2.getColumn(2);
						String[] destinationMailboxes = { "Mailbox", "Archive Folder", "Public Folder" };
						JComboBox<String> jcomboxmail = new JComboBox<String>(destinationMailboxes);
						testColumn.setCellEditor(new DefaultCellEditor(jcomboxmail));
						for (int i = 0; i < table_fileselect.getRowCount(); i++) {

							dtm.addRow(new Object[] { table_fileselect.getValueAt(i, 0).toString(),
									table_fileselect.getValueAt(i, 1).toString(), "Mailbox" });

						}

						return null;
					}

					@Override
					protected void done() {
//				
						obTh.close();
					}
				};

				sw1.execute();
			}
		});
		button.setFocusable(false);

		button.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				button.setIcon(new ImageIcon(backup.class.getResource("/next-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				button.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));
			}
		});

		button.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));

		button.setFocusTraversalKeysEnabled(false);
		button.setFocusPainted(false);
		button.setDefaultCapable(false);
		// btnNext.setToolTipText("Click here to Go To the tree panel ");
		button.setContentAreaFilled(false);
		button.setBorderPainted(false);
		button.setBounds(927, 567, 142, 38);
		panel_fileselect.add(button);

		JScrollPane scrollPane_6 = new JScrollPane();
		scrollPane_6.setBounds(0, 21, 1065, 498);
		panel_fileselect.add(scrollPane_6);

		Object[][] data2 = {};

		String[] cols2 = { "Filename", "" };

		DefaultTableModel tablemodel1 = new DefaultTableModel(data2, cols2) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {
				// all cells false
				if (column == 1) {
					return true;
				} else {
					return false;
				}
			}

			public Class<?> getColumnClass(int column) {
				if (column == 0) {
					return String.class;
				} else {
					return String.class;
				}
			}
		};

		table_fileselect = new JTable(tablemodel1);
		uservaluescomboBox = new JComboBox<>();
		table_fileselect.setFont(new Font("Calibri", Font.PLAIN, 14));
		table_fileselect.setFillsViewportHeight(true);
		table_fileselect.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		table_fileselect.getTableHeader().setReorderingAllowed(false);
		table_fileselect.getTableHeader().setFont(new Font("Calibri", Font.BOLD, 14));
		table_fileselect.setRowHeight(25);
		DefaultTableCellRenderer tablerenderer1 = (DefaultTableCellRenderer) table_fileselect.getTableHeader()
				.getDefaultRenderer();
		tablerenderer1.setHorizontalAlignment(0);
		TableColumnModel csvcolumnModel1 = table_fileselect.getColumnModel();
		TableColumn testColumn = csvcolumnModel1.getColumn(1);
		testColumn.setCellEditor(new DefaultCellEditor(uservaluescomboBox));
		csvcolumnModel1.getColumn(1).setHeaderRenderer(new ImageRenderer1());
		scrollPane_6.setViewportView(table_fileselect);

		JLabel lblNewLabel_6 = new JLabel("");
		lblNewLabel_6.setIcon(new ImageIcon(backup.class.getResource("/bottom.png")));
		lblNewLabel_6.setBounds(0, 555, 1075, 63);
		panel_fileselect.add(lblNewLabel_6);

		panel_mailfolder = new JPanel();
		panel_mailfolder.setBackground(Color.WHITE);
		Cardlayout.add(panel_mailfolder, "panel_mailfolder");
		panel_mailfolder.setLayout(null);

		JScrollPane scrollPane_7 = new JScrollPane();
		scrollPane_7.setBounds(10, 11, 843, 515);
		panel_mailfolder.add(scrollPane_7);

		Object[][] data3 = {};

		String[] cols3 = { "File Name", "Mailbox User Id", "" };

		DefaultTableModel mailtablemodel = new DefaultTableModel(data3, cols3) {

			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {
				// all cells false
				if (column == 2) {
					return true;
				}
				return false;
			}
		};

		table_2 = new JTable(mailtablemodel);
		table_2.setFont(new Font("Calibri", Font.PLAIN, 14));
		table_2.setFillsViewportHeight(true);
		table_2.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		table_2.getTableHeader().setReorderingAllowed(false);
		table_2.getTableHeader().setFont(new Font("Calibri", Font.BOLD, 14));
		table_2.setRowHeight(25);
		DefaultTableCellRenderer tablerenderer2 = (DefaultTableCellRenderer) table_fileselect.getTableHeader()
				.getDefaultRenderer();
		tablerenderer2.setHorizontalAlignment(0);
		csvcolumnModel2 = table_2.getColumnModel();

		csvcolumnModel2.getColumn(2).setHeaderRenderer(new ImageRenderer1());
		scrollPane_7.setViewportView(table_2);

		JButton button_next2 = new JButton("");
		button_next2.setToolTipText("Click here to Go Forward.");
		button_next2.setRolloverEnabled(false);
		button_next2.setRequestFocusEnabled(false);
		button_next2.setOpaque(false);
		button_next2.setFocusable(false);
		button_next2.setFocusTraversalKeysEnabled(false);
		button_next2.setFocusPainted(false);
		button_next2.setDefaultCapable(false);
		button_next2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button_next2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				button_next2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
			}
		});

		button_next2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
		button_next2.setContentAreaFilled(false);
		button_next2.setBorderPainted(false);
		button_next2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (fileoptionm.equalsIgnoreCase("Mbox")) {
					comboBox_1.setVisible(true);
				}
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_3");

			}
		});
		button_next2.setBounds(927, 567, 142, 38);
		panel_mailfolder.add(button_next2);

		JButton button_3 = new JButton("");
		button_3.setToolTipText("Click here to Go Back.\r\n");
		button_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DefaultTableModel model = (DefaultTableModel) table_2.getModel();
				for (int i = 0; i < table_2.getRowCount(); i++) {
					model.removeRow(i);

				}
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_fileselect");
			}
		});
		button_3.setRequestFocusEnabled(false);
		button_3.setRolloverEnabled(false);
		button_3.setOpaque(false);
		button_3.setFocusable(false);
		button_3.setFocusTraversalKeysEnabled(false);
		button_3.setFocusPainted(false);
		button_3.setDefaultCapable(false);
		button_3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button_3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				button_3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});
		button_3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		;
		button_3.setContentAreaFilled(false);
		button_3.setBorderPainted(false);
		button_3.setBounds(800, 567, 142, 38);
		panel_mailfolder.add(button_3);

		JLabel lblNewLabel_7 = new JLabel("");
		lblNewLabel_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		lblNewLabel_7.setBounds(0, 548, 1075, 70);
		panel_mailfolder.add(lblNewLabel_7);

		JPanel panel_9 = new JPanel();
		panel_9.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Select Restore Options",
				TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_9.setBackground(Color.WHITE);
		panel_9.setBounds(863, 11, 202, 515);
		panel_mailfolder.add(panel_9);
		panel_9.setLayout(null);

		JEditorPane dtrpnIfYouWant = new JEditorPane();
		dtrpnIfYouWant.setEditable(false);
		dtrpnIfYouWant.setForeground(Color.WHITE);
		dtrpnIfYouWant.setBackground(Color.DARK_GRAY);
		dtrpnIfYouWant.setFont(new Font("Arial", Font.PLAIN, 14));
		dtrpnIfYouWant.setText("If  you  want  to  take  the   backup  of Mailbox " + System.lineSeparator()
				+ ", Public Folder or " + System.lineSeparator() + "Archive Folder, then" + System.lineSeparator()
				+ "Select from above given " + System.lineSeparator() + "options.");
		dtrpnIfYouWant.setBounds(10, 356, 182, 150);
		panel_9.add(dtrpnIfYouWant);

		JLabel lblNotes = new JLabel("NOTES :");
		lblNotes.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNotes.setBounds(10, 318, 78, 27);
		panel_9.add(lblNotes);

		JRadioButton rdbtnMailbox = new JRadioButton("Mailbox");
		rdbtnMailbox.setSelected(true);
		rdbtnMailbox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_2.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Mailbox", i, 2);
				}
			}
		});
		buttonGroup.add(rdbtnMailbox);
		rdbtnMailbox.setForeground(Color.BLACK);
		rdbtnMailbox.setFont(new Font("Arial", Font.PLAIN, 16));
		rdbtnMailbox.setBackground(Color.WHITE);
		rdbtnMailbox.setBounds(36, 39, 129, 27);
		panel_9.add(rdbtnMailbox);

		JRadioButton rdbtnArchieve = new JRadioButton("In-Place Archive");
		rdbtnArchieve.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_2.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Archive Folder", i, 2);
				}
			}
		});
		buttonGroup.add(rdbtnArchieve);
		rdbtnArchieve.setForeground(Color.BLACK);
		rdbtnArchieve.setFont(new Font("Arial", Font.PLAIN, 16));
		rdbtnArchieve.setBackground(Color.WHITE);
		rdbtnArchieve.setBounds(34, 85, 158, 27);
		panel_9.add(rdbtnArchieve);

		JRadioButton rdbtnPublicFolder = new JRadioButton("Public Folder");
		rdbtnPublicFolder.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_2.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Public Folder", i, 2);
				}
			}
		});
		buttonGroup.add(rdbtnPublicFolder);
		rdbtnPublicFolder.setForeground(Color.BLACK);
		rdbtnPublicFolder.setFont(new Font("Arial", Font.PLAIN, 16));
		rdbtnPublicFolder.setBackground(Color.WHITE);
		rdbtnPublicFolder.setBounds(34, 133, 129, 27);
		panel_9.add(rdbtnPublicFolder);

		panel_3 = new JPanel();
		panel_3.setBackground(Color.WHITE);
		Cardlayout.add(panel_3, "panel_3");
		panel_3.setLayout(null);
		final String sop;
		if (fileoptionm.equalsIgnoreCase("Exchange Backup & Restore")) {
			sop = "Live Exchange";
		} else {
			sop = "OFFICE 365";

		}
		comboBox_fileDestination_type = new JComboBox(new String[] { sop });

		comboBox_fileDestination_type.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				panel_3_.setVisible(false);

				panel_3_2.setVisible(false);

				panel_3_1_2.setVisible(false);

				panel_3_1_2.setVisible(false);

				panel_3_1_2_1.setVisible(false);

				textField_domain_name_p3.setText("");
				output = false;
				tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
				textField_username_p3.setText("");

				passwordField_p3.setText("");
				lbl_progressreport.setText("");
				btn_converter_1.setEnabled(false);
				panel_progress.setVisible(true);
				textField.setVisible(false);
				lblPortNo.setVisible(false);
				comboBox.setVisible(false);
				btn_signout_p3.setVisible(false);

				lblMakeSureYou.setVisible(true);
				lblEnableImap_p3.setVisible(true);
				lblTurnOffTwo_p3.setVisible(true);
				lblAllowLessSecure_p3.setVisible(true);
				panel_5.setVisible(false);

				task_box.setVisible(true);
				panel_taskfilter.setVisible(true);

				if (arg0.getSource() == comboBox_fileDestination_type) {

					JComboBox cb = (JComboBox) arg0.getSource();

					filetype = (String) cb.getSelectedItem();

				}

				panel_progress.setVisible(true);
				panel_3_.setVisible(true);

				CardLayout card = (CardLayout) panel_3_.getLayout();
				card.show(panel_3_, "panel_3_1_1");

				btn_converter_1.setVisible(true);
				btn_converter_1.setEnabled(true);

				if (!(filetype.equalsIgnoreCase("PST") || filetype.equalsIgnoreCase("Thunderbird")
						|| filetype.equalsIgnoreCase("Opera Mail") || filetype.equalsIgnoreCase("OST")
						|| filetype.equalsIgnoreCase("MBOX") || filetype.equalsIgnoreCase("CSV"))) {
					comboBox.setVisible(true);
					panel_5.setVisible(true);
				}
				if (fileoptionm.equalsIgnoreCase("EML Files (.eml)") || filetype.equalsIgnoreCase("Thunderbird")
						|| filetype.equalsIgnoreCase("Opera Mail") || fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")
						|| fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {

					task_box.setVisible(false);
					panel_taskfilter.setVisible(false);

				}

			}
		});
		comboBox_fileDestination_type.setBounds(279, 9, 463, 29);
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
				comboBox_fileDestination_type.setSelectedItem(sop);
			}
		});
		comboBox_fileDestination_type.setVisible(false);
		panel_3.add(comboBox_fileDestination_type);

		btn_previous_p3 = new JButton("");
		btn_previous_p3.setToolTipText("Click here to Go Back.");
		btn_previous_p3.setBorderPainted(false);
		btn_previous_p3.setContentAreaFilled(false);
		btn_previous_p3.setDefaultCapable(false);
		btn_previous_p3.setFocusable(false);
		btn_previous_p3.setFocusTraversalKeysEnabled(false);
		btn_previous_p3.setFocusPainted(false);
		btn_previous_p3.setRolloverEnabled(false);
		btn_previous_p3.setRequestFocusEnabled(false);
		btn_previous_p3.setOpaque(false);
		btn_previous_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				hm.clear();
				lblNew_setemail.setText("");

				lblNew_setsubject.setText("");
				textField_customfolder.setEnabled(true);
				chckbxCustomFolderName.setEnabled(true);
				chckbxrestoretodefault.setEnabled(true);
				label_date.setText("");
				editorPane.setText("");

				try {
					DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

					while (model.getRowCount() > 0) {

						for (int i = 0; i < model.getRowCount(); ++i) {

							model.removeRow(i);
						}
					}

					DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

					while (model1.getRowCount() > 0) {

						for (int i = 0; i < model1.getRowCount(); ++i) {

							model1.removeRow(i);
						}
					}

					if (output) {
						if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
								|| filetype.equals("Hotmail")) {
							clientforexchange_output.dispose();

						} else {

							iconnforimap_output.dispose();
						}
					}
				} catch (Exception e1) {

				}
				comboBox_fileDestination_type.setEnabled(true);
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_mailfolder");
			}
		});
		btn_previous_p3.setBounds(800, 567, 123, 38);
		panel_3.add(btn_previous_p3);
		btn_previous_p3.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_signout_p3 = new JButton("");
		btn_signout_p3.setToolTipText("Click here to Sign Out.");
		btn_signout_p3.setFocusable(false);
		btn_signout_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
			}
		});

		btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
		btn_signout_p3.setFocusTraversalKeysEnabled(false);
		btn_signout_p3.setFocusPainted(false);
		btn_signout_p3.setDefaultCapable(false);
		btn_signout_p3.setContentAreaFilled(false);
		btn_signout_p3.setBorderPainted(false);
		btn_signout_p3.setRolloverEnabled(false);
		btn_signout_p3.setRequestFocusEnabled(false);
		btn_signout_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					String warn = "Do you want to sign out?";
					int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
					if (ans == JOptionPane.YES_OPTION) {

						if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
								|| filetype.equals("Hotmail")) {
							clientforexchange_output.dispose();

						} else {

							iconnforimap_output.dispose();
						}

						textField_domain_name_p3.setText("");
						passwordField_p3.setText("");
						textField_domain_name_p3.setText("");

						btn_converter_1.setEnabled(false);

						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_2");
					}

				} catch (Error e) {
					mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
					return;
				} finally {
					comboBox_fileDestination_type.setEnabled(true);
				}
			}
		});
		btn_signout_p3.setBounds(870, 5, 142, 38);
		btn_signout_p3.setVisible(false);
		panel_3.add(btn_signout_p3);

		btn_converter_1 = new JButton("");
		btn_converter_1.setToolTipText("Click here to Start Process.");
		btn_converter_1.setEnabled(false);

		btn_converter_1.setRequestFocusEnabled(false);
		btn_converter_1.setOpaque(false);
		btn_converter_1.setFocusTraversalKeysEnabled(false);
		btn_converter_1.setFocusable(false);
		btn_converter_1.setFocusPainted(false);
		btn_converter_1.setContentAreaFilled(false);
		btn_converter_1.setBorderPainted(false);
		btn_converter_1.setDefaultCapable(false);
		btn_converter_1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_converter_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_converter_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
			}
		});

		btn_converter_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
		btn_converter_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				count_destination = 0;

				if (chckbx_calender_box.isSelected()) {
					try {
						Calendar calendarstartdate = dateChooser_calender_start.getCalendar();
						Calendar calendarenddate = dateChooser_calendar_end.getCalendar();
						calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
						calendarstartdate.set(Calendar.MINUTE, 00);
						calendarstartdate.set(Calendar.SECOND, 00);
						calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
						calendarenddate.set(Calendar.MINUTE, 59);
						calendarenddate.set(Calendar.SECOND, 59);

						Calenderfilterstartdate = calendarstartdate.getTime();
						Calenderfilterenddate = calendarenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in Calendar filter before Continuing",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}
				}
				if (chckbx_Mail_Filter.isSelected()) {
					try {
						Calendar mailstartdate = dateChooser_mail_fromdate.getCalendar();
						Calendar mailenddate = dateChooser_mail_tilldate.getCalendar();
						mailstartdate.set(Calendar.HOUR_OF_DAY, 00);
						mailstartdate.set(Calendar.MINUTE, 00);
						mailstartdate.set(Calendar.SECOND, 00);
						mailenddate.set(Calendar.HOUR_OF_DAY, 23);
						mailenddate.set(Calendar.MINUTE, 59);
						mailenddate.set(Calendar.SECOND, 59);

						mailfilterstartdate = mailstartdate.getTime();
						mailfilterenddate = mailenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in Mail filter before Continuing",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}
				}
				if (task_box.isSelected()) {
					try {

						Calendar taskstartdate = dateChooser_task_start_date.getCalendar();
						Calendar taskenddate = dateChooser_task_end_date.getCalendar();
						taskstartdate.set(Calendar.HOUR_OF_DAY, 00);
						taskstartdate.set(Calendar.MINUTE, 00);
						taskstartdate.set(Calendar.SECOND, 00);
						taskenddate.set(Calendar.HOUR_OF_DAY, 23);
						taskenddate.set(Calendar.MINUTE, 59);
						taskenddate.set(Calendar.SECOND, 59);

						taskfilterstartdate = taskstartdate.getTime();
						taskfilterenddate = taskenddate.getTime();
					} catch (Exception e1) {
						JOptionPane.showMessageDialog(mf, "Please Enter the date in task filter before Continuing",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}

				}
				if (chckbxCustomFolderName.isSelected() && textField_customfolder.getText().isEmpty()) {

					JOptionPane.showMessageDialog(mf, "Please Enter the name of folder before Continuing",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					throw new NullPointerException();

				}
				Desktop desktop = Desktop.getDesktop();
				th = new Thread(new Runnable() {

					@Override
					public void run() {
						try {
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							comboBox_1.setEnabled(false);
							btnStop.setVisible(true);
							comboBox_fileDestination_type.setEnabled(false);
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							lbl_progressreport.setText("");
							chckbxrestoretodefault.setEnabled(false);
							dateChooser_calender_start.setEnabled(false);
							chckbxRemoveDuplicacy.setEnabled(false);
							dateChooser_calendar_end.setEnabled(false);
							dateChooser_mail_fromdate.setEnabled(false);
							dateChooser_mail_tilldate.setEnabled(false);
							dateChooser_task_start_date.setEnabled(false);
							dateChooser_task_end_date.setEnabled(false);
							btn_signout_p3.setVisible(false);
							panel_5.setVisible(false);
							btn_Destination.setEnabled(false);
							btn_previous_p3.setEnabled(false);
							chckbxCustomFolderName.setEnabled(false);
							textField_customfolder.setEnabled(false);
							btn_converter_1.setEnabled(false);
							btnStop.setVisible(true);
							lbl_progressreport.setText("");
							chckbx_Mail_Filter.setEnabled(false);
							chckbx_calender_box.setEnabled(false);
							comboBox.setVisible(false);

							task_box.setEnabled(false);
							long starttime = System.currentTimeMillis();

							if (checkconvertagain) {

								if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
										|| filetype.equalsIgnoreCase("AOL")
										|| filetype.equalsIgnoreCase("Live Exchange ")
										|| filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Hotmail") || filetype.equalsIgnoreCase("IMAP"))) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ filetype);
										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();
										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")" + filetype);
											f.mkdirs();

											destination_path = f.getAbsolutePath();

										}

									} else {

										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname + filetype);
										f.mkdirs();
										destination_path = tf_Destination_Location.getText() + File.separator
												+ calendertime + File.separator + fname + filetype;
									}

								}

							} else {

								if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
										|| filetype.equalsIgnoreCase("AOL")
										|| filetype.equalsIgnoreCase("Live Exchange ")
										|| filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Hotmail") || filetype.equalsIgnoreCase("IMAP"))) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(
												tf_Destination_Location.getText() + File.separator + customerfolder);
										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();
										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")");
											f.mkdirs();

											destination_path = f.getAbsolutePath();

										}

									} else {

										fname = fileoptionm;

										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname);
										f.mkdirs();
										destination_path = f.getAbsolutePath();
									}
								}

							}

							if (chckbx_Mail_Filter.isSelected()
									&& (mailfilterenddate == null || mailfilterstartdate == null)) {
								JOptionPane.showMessageDialog(mf, "Please Select Start and End Date", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else if (chckbx_calender_box.isSelected()
									&& (Calenderfilterenddate == null || Calenderfilterstartdate == null)) {
								JOptionPane.showMessageDialog(mf, "Please Select Start and End Date", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else {
								String finalpath = destination_path;
								Progressbar.setVisible(true);
								long totalcount = 0;
								mf.logger.info("Convertion into " + fileoptionm + "Start Time : " + cal.getTime()
										+ System.lineSeparator());
								Boolean flagcheck = true;
								for (int i = 0; i < filesfin.length; i++) {
									if (stop) {
										break;
									}

									count_destination = 0;
									file = new File(filesfin[i].replace("<html><b>", ""));
									destination_path = finalpath;

									fname = file.getName();

									mf.logger.info("file name " + fname + "Start Time : " + cal.getTime()
											+ System.lineSeparator());
									if (!(filetype.equalsIgnoreCase("gmail") || filetype.equalsIgnoreCase("YAHOO MAIL")
											|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("IMAP"))) {
										path = "";
									}

									if (table.getValueAt(i, 3).toString().replace("<html><b>", "")
											.equalsIgnoreCase("File")) {

										for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
											if (stop) {
												break;
											}
											if (entry.getKey().trim()
													.equalsIgnoreCase(filesfin[i].replace("<html><b>", ""))) {

												if (!(fileoptionm.equalsIgnoreCase("MBOX")
														|| fileoptionm.equalsIgnoreCase("EML Files (.eml)")
														|| fileoptionm.equalsIgnoreCase("Message Files (.msg)")
														|| fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)"))) {
													pstfolderlist = new ArrayList<String>();

													pstfolderlist.addAll(entry.getValue());
												}

												filepath = file.getAbsolutePath();
												// System.out.println(filepath);

												fname = file.getName().replace(".mbx", "").replace(".mbox", "")
														.replace(".pst", "").replace(".ost", "").replace(".nsf", "")
														.replace(".eml", "").replace(".olm", "");
												fname = getRidOfIllegalFileNameCharacters(fname);
												fname = fname.trim();
												destination_path = destination_path + File.separator + fname;

												label_11.setIcon(
														new ImageIcon(Main_Frame.class.getResource("/download.png")));

												String user = "";
												String folde = "";

												for (int j = 0; j < table_2.getRowCount(); j++) {

													if (table_2.getValueAt(j, 0).toString().equals(file.getName())) {
														user = table_2.getValueAt(j, 1).toString();
														folde = table_2.getValueAt(j, 2).toString();
														break;
													}

												}
												try {
													clientforexchange_output = EWSClient.getEWSClient(mailboxUri,
															username_p3, password_p3);
													clientforexchange_output.getMailboxInfo(user);
												} catch (Exception e1) {

													clientforexchange_output = EWSClient.getEWSClient(mailboxUri,
															username_p3, password_p3);

													clientforexchange_output.impersonateUser(3, user);
												}

												String Publicfolderuri = "";
												if (folde.equalsIgnoreCase("Public Folder")) {
													ExchangeFolderInfo folderj = clientforexchange_output
															.getFolderInfo("PublicFoldersRoot");

													Publicfolderuri = folderj.getUri();
												}

												if (chckbxCustomFolderName.isSelected()) {
													String customerfolder = textField_customfolder.getText()
															.replace("//s", "");

													customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
													try {
														if (folde.equalsIgnoreCase("Archive Folder")) {

															Folderuri = clientforexchange_output
																	.createFolder(customerfolder).getUri();
															fa = Folderuri;
														} else if (folde.equalsIgnoreCase("Public Folder")) {
															try {
																Folderuri = clientforexchange_output
																		.createFolder(Publicfolderuri, customerfolder)
																		.getUri();
																fa = Folderuri;
															} catch (Exception e) {
																JOptionPane.showMessageDialog(restore.this,
																		"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder is created to perform this operation  on this object. See the folder contact or your system administrator",
																		messageboxtitle, JOptionPane.ERROR_MESSAGE,
																		new ImageIcon(backup.class
																				.getResource("/information.png")));
															}
														} else {
															Folderuri = clientforexchange_output
																	.createFolder(customerfolder).getUri();
															fa = Folderuri;
														}

													} catch (Exception e) {

														Folderuri = clientforexchange_output
																.createFolder(customerfolder + "(" + calendertime + ")")
																.getUri();
														fa = Folderuri;

													}
												} else {

													if (folde.equalsIgnoreCase("Public Folder")) {
														try {
															Folderuri = clientforexchange_output
																	.createFolder(Publicfolderuri, calendertime)
																	.getUri();
															fa = Folderuri;
														} catch (Exception e) {
															JOptionPane.showMessageDialog(restore.this,
																	"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder  is created to perform this operation  on this object. See the folder contact or your system administrator",
																	messageboxtitle, JOptionPane.ERROR_MESSAGE,
																	new ImageIcon(backup.class
																			.getResource("/information.png")));
														}
													} else {
														if (flagcheck) {
															Folderuri = clientforexchange_output
																	.createFolder(calendertime).getUri();
															flagcheck = false;
															fa = Folderuri;
														}
													}
												}
												if (!folde.equalsIgnoreCase("Archive Folder")) {

													if (chckbxrestoretodefault.isSelected()) {
														try {

															if (fileoptionm.equalsIgnoreCase("Mbox")) {
																ConvertMbox_defaultrestore();
															} else {
																ConvertPST_defaultrestore();
															}

														} catch (Exception e) {

														}

													} else {
														Folderuri = clientforexchange_output
																.createFolder(Folderuri, fname).getUri();
														try {

															if (fileoptionm.equalsIgnoreCase("Mbox")) {
																ConvertMbox_365(folde);
															} else {
																ConvertPSTOST_365(folde);
															}

														} catch (Exception e) {

														}

													}
												} else {
													if (chckbxrestoretodefault.isSelected()) {
														JOptionPane.showMessageDialog(mf,
																"As you have selected Archive Folder as well as selected Default folder it will be stored in Archive Folder",
																messageboxtitle, JOptionPane.ERROR_MESSAGE,
																new ImageIcon(Main_Frame.class
																		.getResource("/information.png")));
													}
													Folderuri = clientforexchange_output.createFolder(Folderuri, fname)
															.getUri();

													try {

														if (fileoptionm.equalsIgnoreCase("Mbox")) {
															ConvertMbox_365(folde);
														} else {
															ConvertPSTOST_365(folde);
														}

													} catch (Exception e) {

													}

												}
												Folderuri = fa;

											}
										}
									} else if (table.getValueAt(i, 3).toString().replace("<html><b>", "")
											.equalsIgnoreCase("Folder")) {

										parent = file.getParent();

										// progressBar_message_p3.setVisible(true);

										label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));

										try {
											ConvertPSTOST_folder_exchange(file);
										} catch (Exception e) {

										}

										Folderuri = fa;

									}

									if (filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("GMAIL")
											|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equals("AOL")
											|| filetype.equalsIgnoreCase("Hotmail")
											|| filetype.equalsIgnoreCase("Live Exchange ")
											|| filetype.equalsIgnoreCase("IMAP")) {

										if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
											destination_path = "http://login.yahoo.com";
										} else if (filetype.equalsIgnoreCase("GMAIL")) {
											destination_path = "https://mail.google.com";
										} else if (filetype.equals("AOL")) {
											destination_path = "https://login.aol.com";
										} else if (filetype.equals("IMAP")) {
											destination_path = "IMAP";
										} else if (filetype.equals("Hotmail")) {
											destination_path = "https://outlook.live.com";
										} else if (filetype.equals("Live Exchange")) {
											destination_path = "Live Exchange";
										} else {
											destination_path = "https://outlook.office365.com";
										}
									}
									String duration = Duration(starttime).toString();
									mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
									mode.addRow(new Object[] { fileoptionm, filetype, fname, Status, duration,
											count_destination, destination_path });
									destination_path = destination_path.replace(File.separator + fname, "");
									totalcount = totalcount + count_destination;
									mf.logger.info("File Saved " + count_destination + System.lineSeparator()
											+ "End Time : " + cal.getTime() + System.lineSeparator()
											+ "**********************************************************");

								}

								mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();

								mode.addRow(new Object[] { "Total Message", "", "", "", "", totalcount, "" });
								mf.logger.info("File Saved " + totalcount + System.lineSeparator() + "End Time : "
										+ cal.getTime() + System.lineSeparator()
										+ "*****************************END*****************************");

							}

							destination_path = destination_path.replace(File.separator + fname, "");
							if (filetype.equalsIgnoreCase("THUNDERBIRD")) {
								JOptionPane.showMessageDialog(mf,
										"Please open the converted file from " + destination_path + " Thunderbird",
										messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
							}
							if (filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("GMAIL")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equals("AOL")
									|| filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("hotmail") || filetype.equalsIgnoreCase("IMAP")) {
								if (Desktop.isDesktopSupported()
										&& Desktop.getDesktop().isSupported(Desktop.Action.BROWSE)) {

									if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

										reportpath = "http://login.yahoo.com";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("GMAIL")) {

										reportpath = "https://mail.google.com";
										openBrowser(reportpath);
									} else if (filetype.equals("AOL")) {

										reportpath = "https://login.aol.com";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("Live Exchange")) {

										JOptionPane.showMessageDialog(mf,
												"Please open the converted file from Live Exchange", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("Hotmail")) {

										reportpath = "https://outlook.live.com";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("IMAP")) {

										JOptionPane.showMessageDialog(mf, "Please open the converted file from  IMAP",
												messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
									} else {

										reportpath = "https://outlook.office365.com";
										openBrowser(reportpath);
									}
								}

							} else {

								desktop.open(f);
								reportpath = f.getAbsolutePath();

							}

						} catch (Exception e) {

						} finally {
							Progressbar.setVisible(false);
							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_4");
							count_eml_msg_emlx = 0;

							try {
								if (filetype.equalsIgnoreCase("MBOX")) {
									if (!(wr == null)) {
										wr.dispose();
									}
								} else if (filetype.equalsIgnoreCase("OST") || filetype.equalsIgnoreCase("PST")) {
									if (!(pst == null)) {
										pst.dispose();

									}

								} else if (filetype.equalsIgnoreCase("CSV")) {
									if (!(writer == null)) {
										try {
											writer.close();
										} catch (IOException e) {

											e.printStackTrace();
										}
									}
								}
							} catch (Exception e) {

							}
							checky = true;
						}
					}

				});

				th.start();

			}
		});

		btn_converter_1.setBounds(933, 571, 130, 30);
		panel_3.add(btn_converter_1);
		btn_converter_1.setFont(new Font("Tahoma", Font.BOLD, 12));

		panel_3_2 = new JPanel();
		panel_3_2.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_2.setBackground(Color.WHITE);
		panel_3_2.setBounds(173, 392, 1021, 57);
		panel_3.add(panel_3_2, "panel_3_2");
		panel_3_2.setLayout(null);
		panel_3_2.setVisible(false);

		tf_Destination_Location = new JTextField();
		tf_Destination_Location.setBounds(64, 24, 785, 20);
		panel_3_2.add(tf_Destination_Location);
		tf_Destination_Location.setEditable(false);
		tf_Destination_Location.setColumns(10);

		btn_Destination = new JButton("");
		btn_Destination.setToolTipText("Click here to select the Destination path.");
		btn_Destination.setRolloverEnabled(false);
		btn_Destination.setRequestFocusEnabled(false);
		btn_Destination.setOpaque(false);
		btn_Destination.setFocusable(false);
		btn_Destination.setFocusTraversalKeysEnabled(false);
		btn_Destination.setFocusPainted(false);
		btn_Destination.setDefaultCapable(false);
		btn_Destination.setContentAreaFilled(false);
		btn_Destination.setBorderPainted(false);
		btn_Destination.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-save-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
			}
		});

		btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
		btn_Destination.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					destinationPath();
				} catch (Exception e1) {

					e1.printStackTrace();
				}

			}
		});
		btn_Destination.setBounds(870, 11, 133, 38);
		panel_3_2.add(btn_Destination);
		btn_Destination.setFont(new Font("Tahoma", Font.BOLD, 12));

		progressBar_message_p3 = new JProgressBar();
		progressBar_message_p3.setBackground(Color.WHITE);
		progressBar_message_p3.setBounds(10, 464, 6, 7);

		panel_3_ = new JPanel();
		panel_3_.setBounds(20, 49, 1019, 304);
		panel_3.add(panel_3_);
		panel_3_.setLayout(new CardLayout(0, 0));

		panel_3_1_1 = new JPanel();
		panel_3_1_1.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_1_1.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_1, "panel_3_1_1");
		panel_3_1_1.setLayout(null);

		JPanel panel_mailfilter = new JPanel();
		panel_mailfilter.setBackground(Color.WHITE);
		panel_mailfilter.setBounds(475, 34, 534, 47);
		panel_3_1_1.add(panel_mailfilter);
		panel_mailfilter.setLayout(null);

		dateChooser_mail_fromdate = new JDateChooser();
		dateChooser_mail_fromdate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_fromdate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_mail_fromdate.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_mail_fromdate.setBounds(130, 16, 120, 22);
		dateChooser_mail_fromdate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_fromdate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_fromdate);

		dateChooser_mail_tilldate = new JDateChooser();
		dateChooser_mail_tilldate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_tilldate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_mail_tilldate.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_mail_tilldate.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_mail_tilldate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_tilldate.setEnabled(false);
		dateChooser_mail_tilldate.setBounds(398, 19, 125, 19);
		panel_mailfilter.add(dateChooser_mail_tilldate);

		JLabel label_3 = new JLabel("End Date");
		label_3.setBounds(317, 16, 71, 19);
		panel_mailfilter.add(label_3);
		label_3.setFont(new Font("Tahoma", Font.BOLD, 15));

		JLabel label_1 = new JLabel("Start Date");
		label_1.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_1.setBounds(21, 16, 82, 20);
		panel_mailfilter.add(label_1);
		label_3.setFont(new Font("Tahoma", Font.BOLD, 15));

		JPanel panel_checkboxcalenderfilter = new JPanel();
		panel_checkboxcalenderfilter.setBackground(Color.WHITE);
		panel_checkboxcalenderfilter.setBounds(12, 86, 119, 36);
		// panel_3_1_1.add(panel_checkboxcalenderfilter);
		panel_checkboxcalenderfilter.setLayout(null);

		JPanel panel_Calender = new JPanel();
		panel_Calender.setBackground(Color.WHITE);
		panel_Calender.setBounds(12, 119, 1050, 53);
		// panel_3_1_1.add(panel_Calender);
		panel_Calender.setLayout(null);

		chckbx_calender_box = new JCheckBox("Calendar Filter");
		chckbx_calender_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_Calender.setEnabled(true);
					dateChooser_calendar_end.setEnabled(true);
					dateChooser_calender_start.setEnabled(true);
				} else {
					panel_Calender.setEnabled(false);
					dateChooser_calendar_end.setEnabled(false);
					dateChooser_calender_start.setEnabled(false);
				}
			}
		});
		chckbx_calender_box.setBackground(Color.WHITE);
		chckbx_calender_box.setBounds(0, 9, 113, 25);
		panel_checkboxcalenderfilter.add(chckbx_calender_box);

		dateChooser_calender_start = new JDateChooser();
		dateChooser_calender_start.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calender_start.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_calender_start.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_calender_start.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_calender_start.setEnabled(false);
		dateChooser_calender_start.setBounds(101, 25, 168, 20);
		panel_Calender.add(dateChooser_calender_start);

		dateChooser_calendar_end = new JDateChooser();
		dateChooser_calendar_end.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calendar_end.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_calendar_end.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_calendar_end.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_calendar_end.setBounds(893, 25, 147, 19);
		panel_Calender.add(dateChooser_calendar_end);
		dateChooser_calendar_end.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_calendar_end.setEnabled(false);

		JLabel label = new JLabel("End Date");
		label.setBounds(815, 25, 79, 19);
		panel_Calender.add(label);
		label.setFont(new Font("Tahoma", Font.BOLD, 15));

		JLabel label_2 = new JLabel("Start Date");
		label_2.setBounds(12, 25, 91, 20);
		panel_Calender.add(label_2);
		label_2.setFont(new Font("Tahoma", Font.BOLD, 15));

		panel_taskfilter = new JPanel();
		panel_taskfilter.setBackground(Color.WHITE);
		panel_taskfilter.setBounds(475, 121, 534, 47);
		panel_3_1_1.add(panel_taskfilter);
		panel_taskfilter.setLayout(null);
		JLabel label_8 = new JLabel("Start Date");
		label_8.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_8.setBounds(21, 14, 88, 20);
		panel_taskfilter.add(label_8);

		JLabel label_9 = new JLabel("End Date");
		label_9.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_9.setBounds(316, 15, 75, 19);
		panel_taskfilter.add(label_9);

		dateChooser_task_start_date = new JDateChooser();
		dateChooser_task_start_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_start_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();
				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_task_start_date.setMaxSelectableDate(startdate);

			}
		});
		dateChooser_task_start_date.setEnabled(false);
		dateChooser_task_start_date.setBounds(131, 14, 119, 22);
		panel_taskfilter.add(dateChooser_task_start_date);

		dateChooser_task_end_date = new JDateChooser();
		dateChooser_task_end_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_start_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_task_end_date.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_task_end_date.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Exception e1) {
					return;
				}
			}
		});
		dateChooser_task_end_date.setEnabled(false);
		dateChooser_task_end_date.setBounds(399, 14, 125, 19);
		panel_taskfilter.add(dateChooser_task_end_date);

		task_box = new JCheckBox("Task Filter");
		task_box.setBounds(475, 98, 75, 16);
		panel_3_1_1.add(task_box);
		task_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					dateChooser_task_end_date.setEnabled(true);
					dateChooser_task_start_date.setEnabled(true);
				} else {

					dateChooser_task_end_date.setEnabled(false);
					dateChooser_task_start_date.setEnabled(false);
				}
			}
		});
		task_box.setBackground(Color.WHITE);

		chckbx_Mail_Filter = new JCheckBox("Mail Filter");
		chckbx_Mail_Filter.setBounds(475, 7, 113, 25);
		panel_3_1_1.add(chckbx_Mail_Filter);
		chckbx_Mail_Filter.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_mailfilter.setEnabled(true);
					dateChooser_mail_fromdate.setEnabled(true);
					dateChooser_mail_tilldate.setEnabled(true);
				}

				else {
					panel_mailfilter.setEnabled(false);
					dateChooser_mail_fromdate.setEnabled(false);
					dateChooser_mail_tilldate.setEnabled(false);
				}

			}
		});

		chckbx_Mail_Filter.setBackground(Color.WHITE);

		panel_5 = new JPanel();
		panel_5.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_5.setBackground(Color.WHITE);
		panel_5.setBounds(10, 179, 527, 38);
		panel_5.setVisible(false);

		panel_5.setLayout(null);

		comboBox = new JComboBox<String>();
		comboBox.setBounds(192, 11, 294, 16);
		panel_5.add(comboBox);
		comboBox.addItem("Subject");

		comboBox.addItem("Subject_Date");
		comboBox.addItem("Date_Subject");
		comboBox.addItem("From_Subject_Date");
		comboBox.addItem("Date_From_Subject");
		comboBox.setVisible(false);

		panel_3_1_2 = new JPanel();
		panel_3_1_2.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_1_2.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_2, "panel_3_1_2");
		panel_3_1_2.setLayout(null);

		lblLive_Chat_p3 = new JLabel("More Help");
		lblLive_Chat_p3.setForeground(Color.RED);
		lblLive_Chat_p3.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblLive_Chat_p3.setBounds(818, 16, 91, 25);
		lblLive_Chat_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");
			}
		});

		panel_3_1_2.add(lblLive_Chat_p3);

		lbl_connecting_p3 = new JLabel("");
		lbl_connecting_p3.setBounds(437, 239, 85, 32);
		panel_3_1_2.add(lbl_connecting_p3);
		lbl_connecting_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));
		lbl_connecting_p3.setVisible(false);

		JLabel lblNewLabel = new JLabel("User Name");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel.setBounds(29, 16, 102, 25);
		panel_3_1_2.add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel("Password ");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel_1.setBounds(29, 48, 85, 25);
		panel_3_1_2.add(lblNewLabel_1);

		textField_username_p3 = new JTextField();
		textField_username_p3.setComponentPopupMenu(menu);
		textField_username_p3.setBounds(303, 13, 408, 25);
		panel_3_1_2.add(textField_username_p3);
		textField_username_p3.setColumns(10);

		passwordField_p3 = new JPasswordField();
		passwordField_p3.setComponentPopupMenu(menu);
		passwordField_p3.setBounds(303, 49, 408, 25);
		panel_3_1_2.add(passwordField_p3);

		JCheckBox chckbxShowPassword_p3 = new JCheckBox("Show Password");
		chckbxShowPassword_p3.setBackground(Color.WHITE);
		chckbxShowPassword_p3.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField_p3.setEchoChar((char) 0);
				}

				else {
					passwordField_p3.setEchoChar('*');
				}
			}
		});
		chckbxShowPassword_p3.setFont(new Font("Tahoma", Font.BOLD, 13));
		chckbxShowPassword_p3.setBounds(743, 48, 143, 25);
		panel_3_1_2.add(chckbxShowPassword_p3);

		JButton btn_Sign_p3 = new JButton("");
		btn_Sign_p3.setDefaultCapable(false);
		btn_Sign_p3.setBorderPainted(false);
		btn_Sign_p3.setRolloverEnabled(false);
		btn_Sign_p3.setRequestFocusEnabled(false);
		btn_Sign_p3.setOpaque(false);
		btn_Sign_p3.setFocusable(false);
		btn_Sign_p3.setFocusTraversalKeysEnabled(false);
		btn_Sign_p3.setFocusPainted(false);
		btn_Sign_p3.setContentAreaFilled(false);
		btn_Sign_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));

		btn_Sign_p3.setBounds(851, 233, 123, 38);
		panel_3_1_2.add(btn_Sign_p3);
		btn_Sign_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				try {
					domain_p3 = textField_domain_name_p3.getText().replaceAll("//s", "");
					domain_p3 = domain_p3.trim();
				} catch (Exception a) {
					domain_p3 = "";
				}
				try {
					username_p3 = textField_username_p3.getText().replaceAll("//s", "");
					username_p3 = username_p3.trim();
				} catch (Exception a) {
					username_p3 = "";
				}
				try {
					password_p3 = new String(passwordField_p3.getPassword());
					password_p3 = password_p3.trim();
				} catch (Exception a) {
					password_p3 = "";
				}
				try {
					portnofiletype = Integer.parseInt(textField.getText().replaceAll("//s", ""));
				} catch (Exception a) {

				}
				chckbxShowPassword_p3.setEnabled(false);
				textField.setEnabled(false);
				comboBox_fileDestination_type.setEnabled(false);
				passwordField_p3.setEnabled(false);
				textField_username_p3.setEnabled(false);
				textField_domain_name_p3.setEnabled(false);
				if (username_p3.equalsIgnoreCase("") || password_p3.equalsIgnoreCase("")) {

					if (username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(mf, "User name and Password fields cannot be empty",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (username_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "User name field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (password_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "Password field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					}

				}
				if (!(isValid(username_p3))) {

					JOptionPane.showMessageDialog(mf, "Please enter a valid username", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));

				} else {

					th = new Thread(new Runnable() {

						@Override
						public void run() {
							lbl_connecting_p3.setVisible(true);

							try {

								btn_previous_p3.setEnabled(false);
								btn_converter_1.setEnabled(false);

								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									conntiontooffice365_output();

								}

								else if (filetype.equalsIgnoreCase("GMAIL")) {
									connectiontogmail_output();

								} else if (filetype.equalsIgnoreCase("Live Exchange ")) {
									connectionwithexchangeserver_output();

								}

								else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
									connectiontoyahoo_output();

								} else if (filetype.equalsIgnoreCase("HOTMAIL")) {
									conntiontohotmail_output();

								} else if (filetype.equalsIgnoreCase("AOL")) {
									connectiontoaol_output();

								} else if (filetype.equalsIgnoreCase("IMAP")) {
									connectiontoimap_output();

								}
								panel_3_1_2.setVisible(false);
								panel_3_1_2.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_1");
								btn_signout_p3.setVisible(true);
								btn_converter_1.setEnabled(true);
								btn_converter_1.setVisible(true);
								output = true;

							} catch (Exception e) {
								if (filetype.equalsIgnoreCase("Gmail")) {
									if (e.getMessage().contains(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(mf,
												"Connection Not Estalished with Gmail please check your Credantial OR Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(mf, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().contains(
											"AE_3_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(mf,
												"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(mf, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(mf, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(mf, "Connection not established", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}

							} finally {
								lbl_connecting_p3.setVisible(false);

								textField.setEnabled(true);

								passwordField_p3.setEnabled(true);
								textField_username_p3.setEnabled(true);
								textField_domain_name_p3.setEnabled(true);
								chckbxShowPassword_p3.setEnabled(true);
								btn_previous_p3.setEnabled(true);

							}

						}
					});
					th.start();

				}
			}
		});
		btn_Sign_p3.setFont(new Font("Tahoma", Font.BOLD, 14));

		panel_3_1_2_1 = new JPanel();
		panel_3_1_2_1.setBackground(Color.WHITE);
		panel_3_1_2_1.setBounds(10, 84, 805, 53);
		panel_3_1_2.add(panel_3_1_2_1);
		panel_3_1_2_1.setVisible(false);
		panel_3_1_2_1.setLayout(null);

		textField_domain_name_p3 = new JTextField();

		textField_domain_name_p3.setComponentPopupMenu(menu);
		textField_domain_name_p3.setBounds(294, 15, 409, 26);
		panel_3_1_2_1.add(textField_domain_name_p3);
		textField_domain_name_p3.setColumns(10);

		lbl_Domain = new JLabel("");
		lbl_Domain.setBounds(20, 15, 271, 27);
		panel_3_1_2_1.add(lbl_Domain);
		lbl_Domain.setFont(new Font("Tahoma", Font.BOLD, 14));

		lblPortNo = new JLabel("PORT NO");

		lblPortNo.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblPortNo.setBounds(29, 148, 158, 32);
		panel_3_1_2.add(lblPortNo);

		textField = new JTextField();
		textField.setComponentPopupMenu(menu);
		textField.setBounds(303, 148, 408, 32);
		textField.setText(Integer.toString(993));
		panel_3_1_2.add(textField);
		textField.setColumns(10);

		panel = new JPanel();
		panel.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel.setBackground(Color.WHITE);
		panel.setBounds(74, 191, 713, 44);
		panel_3_1_2.add(panel);
		panel.setLayout(null);

		lblMakeSureYou = new JLabel("Make sure you have");
		lblMakeSureYou.setBounds(10, 0, 143, 32);
		panel.add(lblMakeSureYou);
		lblMakeSureYou.setFont(new Font("Tahoma", Font.BOLD, 14));

		lblEnableImap_p3 = new JLabel("Enable IMAP");
		lblEnableImap_p3.setBounds(215, 6, 77, 25);
		panel.add(lblEnableImap_p3);
		lblEnableImap_p3.setForeground(SystemColor.textHighlight);
		lblEnableImap_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblAllowLessSecure_p3 = new JLabel("Allow Less Secure Apps");
		lblAllowLessSecure_p3.setBounds(340, 11, 135, 14);
		panel.add(lblAllowLessSecure_p3);
		lblAllowLessSecure_p3.setForeground(SystemColor.textHighlight);
		lblAllowLessSecure_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					openBrowser(
							"https://www.arysontechnologies.com/office365-backup/disable-multi-factor-authentication-office-365.html");
				} else if (filetype.equalsIgnoreCase("IMAP")) {

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					openBrowser("https://www.arysontechnologies.com/gmail-backup/allow-less-secure-apps.html");
				} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
					openBrowser(
							"https://www.arysontechnologies.com/yahoo-backup/how-to-turn-off-two-step-verification-in-yahoo-mail.html");
				} else if (filetype.equalsIgnoreCase("AOL")) {
					openBrowser(
							"https://www.arysontechnologies.com/aol-backup/add-two-step-verification-and-create-aap-pwd-for-aol.html");
				} else if (filetype.equalsIgnoreCase("HOTMAIL")) {
					openBrowser("https://www.arysontechnologies.com/hotmail-backup/create-app-password-hotmail.html");
				} else if (filetype.equalsIgnoreCase("Live Exchange")) {

				}

			}
		});
		lblAllowLessSecure_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo_p3 = new JLabel("Turn Off Two Step Verification");
		lblTurnOffTwo_p3.setBounds(508, 11, 160, 14);
		panel.add(lblTurnOffTwo_p3);
		lblTurnOffTwo_p3.setForeground(SystemColor.textHighlight);
		lblTurnOffTwo_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				if (filetype.equalsIgnoreCase("OFFICE 365")) {

				} else if (filetype.equalsIgnoreCase("IMAP")) {

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					openBrowser("https://www.arysontechnologies.com/gmail-backup/allow-less-secure-apps.html");
				} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
					openBrowser(
							"https://www.arysontechnologies.com/yahoo-backup/how-to-turn-off-two-step-verification-in-yahoo-mail.html");
				} else if (filetype.equalsIgnoreCase("AOL")) {
					openBrowser(
							"https://www.arysontechnologies.com/aol-backup/add-two-step-verification-and-create-aap-pwd-for-aol.html");
				} else if (filetype.equalsIgnoreCase("HOTMAIL")) {

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {

				}

			}
		});
		lblTurnOffTwo_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));
		lblEnableImap_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					openBrowser(
							"https://www.arysontechnologies.com/office365-backup/create-new-password-for-office-365.html");
				} else if (filetype.equalsIgnoreCase("IMAP")) {

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					openBrowser("https://www.arysontechnologies.com/gmail-backup/allow-less-secure-apps.html");
				} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
					openBrowser(
							"https://www.arysontechnologies.com/yahoo-backup/how-to-generate-third-party-app-access-passwords-yahoo.html");
				} else if (filetype.equalsIgnoreCase("AOL")) {

				} else if (filetype.equalsIgnoreCase("HOTMAIL")) {
					openBrowser("https://www.arysontechnologies.com/hotmail-backup/create-app-password-hotmail.html");
				} else if (filetype.equalsIgnoreCase("Live Exchange")) {

				}

			}
		});

		JLabel label_5 = new JLabel("");
		label_5.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_5.setBounds(0, 561, 1075, 63);
		panel_3.add(label_5);

		panel_progress = new JPanel();
		panel_progress.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_progress.setBackground(Color.WHITE);
		panel_progress.setBounds(20, 448, 1022, 102);
		panel_3.add(panel_progress);
		panel_progress.setLayout(null);

		btnStop = new JButton("");
		btnStop.setToolTipText("Click here to Stop the process.");
		btnStop.setBounds(873, 2, 123, 38);
		panel_progress.add(btnStop);
		btnStop.setContentAreaFilled(false);
		btnStop.setBorderPainted(false);
		btnStop.setBackground(Color.WHITE);
		btnStop.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		btnStop.setRolloverEnabled(false);
		btnStop.setVisible(false);
		btnStop.setRequestFocusEnabled(false);
		btnStop.setOpaque(false);
		btnStop.setFocusable(false);
		btnStop.setFocusTraversalKeysEnabled(false);
		btnStop.setFocusPainted(false);
		btnStop.setDefaultCapable(false);
		SwingUtilities.invokeLater(new Runnable() {

			@Override
			public void run() {
				comboBox.setSelectedItem("Subject");

				lblnamingconvention = new JLabel("Naming Convention");
				lblnamingconvention.setFont(new Font("Tahoma", Font.BOLD, 11));
				lblnamingconvention.setBounds(10, 12, 161, 14);
				panel_5.add(lblnamingconvention);

			}
		});

		Progressbar = new JLabel("");
		Progressbar.setBounds(10, 5, 836, 33);
		panel_progress.add(Progressbar);
		Progressbar.setVisible(false);
		Progressbar.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));

		lbl_progressreport = new JLabel("");
		lbl_progressreport.setBounds(10, 49, 892, 29);

		panel_progress.add(lbl_progressreport);

		JPanel panel_duplicacy = new JPanel();
		panel_duplicacy.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_duplicacy.setBackground(Color.WHITE);
		panel_duplicacy.setBounds(10, 11, 455, 260);
		panel_3_1_1.add(panel_duplicacy);
		panel_duplicacy.setLayout(null);

		chckbxRemoveDuplicacy = new JCheckBox("Remove Duplicate Mail On basis of To, From, Subject, Bcc, Body\r\n");
		chckbxRemoveDuplicacy.setForeground(Color.RED);
		chckbxRemoveDuplicacy.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxRemoveDuplicacy.setBackground(Color.WHITE);
		chckbxRemoveDuplicacy.setBounds(6, 7, 391, 23);
		panel_duplicacy.add(chckbxRemoveDuplicacy);

		label_12 = new JLabel("");
		label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_12.setToolTipText("Remove Duplicate Mail On basis of To, From, Subject, Bcc, Body");
		label_12.setBounds(403, 11, 26, 23);
		panel_duplicacy.add(label_12);

		panel_6 = new JPanel();
		panel_6.setBackground(Color.WHITE);
		panel_6.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_6.setBounds(475, 174, 534, 43);
		panel_3_1_1.add(panel_6);
		panel_6.setLayout(null);

		JLabel lblNewLabel_10 = new JLabel("");
		lblNewLabel_10.setToolTipText("Add the name by which the folder will be created " + "\r\n"
				+ " it must not contain these characters :\\?/|*<>\t");
		lblNewLabel_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lblNewLabel_10.setBounds(454, 10, 24, 25);
		panel_6.add(lblNewLabel_10);

		textField_customfolder = new JTextField();
		textField_customfolder.addKeyListener(new KeyAdapter() {
			public void keyReleased(KeyEvent event) {

				String content = textField_customfolder.getText();
				if (content.contains(":") || content.contains(":") || content.contains("\\") || content.contains("?")
						|| content.contains("/") || content.contains("|") || content.contains("*")
						|| content.contains("<") || content.contains(">") || content.contains("\t")
						|| content.contains("//s") || content.contains("\"")) {
					textField_customfolder.setText(getRidOfIllegalFileNameCharacters(content).trim());
				}
			}
		});
		textField_customfolder.setBounds(163, 10, 267, 26);
		textField_customfolder.setEditable(false);
		panel_6.add(textField_customfolder);
		textField_customfolder.setColumns(10);

		chckbxCustomFolderName = new JCheckBox("Custom Folder Name");
		chckbxCustomFolderName.setBounds(6, 10, 141, 23);
		panel_6.add(chckbxCustomFolderName);
		chckbxCustomFolderName.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxCustomFolderName.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					textField_customfolder.setEditable(true);
				}

				else {
					textField_customfolder.setEditable(false);
				}
			}
		});

		chckbxCustomFolderName.setBackground(Color.WHITE);

		panel_7 = new JPanel();
		panel_7.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_7.setBackground(Color.WHITE);
		panel_7.setBounds(475, 228, 534, 43);
		panel_3_1_1.add(panel_7);
		panel_7.setLayout(null);

		chckbxrestoretodefault = new JCheckBox("Restore to Default Folder (Only For MailBox)");
		chckbxrestoretodefault.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (fileoptionm.equalsIgnoreCase("mbox")) {
					if (arg0.getStateChange() == ItemEvent.SELECTED) {
						textField_customfolder.setText("");
						textField_customfolder.setEnabled(false);
						chckbxCustomFolderName.setSelected(false);
						chckbxCustomFolderName.setEnabled(false);
						comboBox_1.setEnabled(true);
					} else {
						textField_customfolder.setEnabled(true);
						chckbxCustomFolderName.setEnabled(true);
						comboBox_1.setEnabled(false);

					}
				}
			}
		});
		chckbxrestoretodefault.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxrestoretodefault.setForeground(Color.RED);
		chckbxrestoretodefault.setBackground(Color.WHITE);
		chckbxrestoretodefault.setBounds(6, 7, 292, 23);
		panel_7.add(chckbxrestoretodefault);

		comboBox_1 = new JComboBox();
		comboBox_1.setEnabled(false);
		comboBox_1.setVisible(false);
		comboBox_1.addItem("Inbox");
		comboBox_1.addItem("Sent Item");
		comboBox_1.addItem("Draft");
		comboBox_1.setBounds(304, 8, 178, 20);
		panel_7.add(comboBox_1);

		label_11 = new JLabel("");
		label_11.setBounds(926, 36, 86, 42);
		panel_progress.add(label_11);

		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process ? ";
				int ans = JOptionPane.showConfirmDialog(restore.this, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					stop = true;

				}

			}
		});

		panel_4 = new JPanel();
		panel_4.setBackground(Color.WHITE);
		Cardlayout.add(panel_4, "panel_4");
		panel_4.setLayout(null);

		JScrollPane scrollPane_table_panel4 = new JScrollPane();
		scrollPane_table_panel4.setBounds(6, 13, 1063, 406);
		panel_4.add(scrollPane_table_panel4);

		table_fileConvertionreport_panel4 = new JTable() {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};
		table_fileConvertionreport_panel4.getTableHeader().setReorderingAllowed(false);
		table_fileConvertionreport_panel4.setModel(new DefaultTableModel(new Object[][] {},

				new String[] { "From", "To", "File Name", "Status", "Duration", "Message count", "Path" }));
		scrollPane_table_panel4.setViewportView(table_fileConvertionreport_panel4);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(2).setPreferredWidth(126);
		btnDowloadReport = new JButton("");

		btnDowloadReport_1 = new JButton("");
		btnDowloadReport_1.setToolTipText("Click here to Download the Report.");
		btnDowloadReport_1.setRolloverEnabled(false);
		btnDowloadReport_1.setRequestFocusEnabled(false);
		btnDowloadReport_1.setOpaque(false);
		btnDowloadReport_1.setFocusable(false);
		btnDowloadReport_1.setFocusTraversalKeysEnabled(false);
		btnDowloadReport_1.setFocusPainted(false);
		btnDowloadReport_1.setDefaultCapable(false);
		btnDowloadReport_1.setContentAreaFilled(false);
		btnDowloadReport_1.setBorderPainted(false);
		btnDowloadReport_1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));
			}
		});

		btnDowloadReport_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));

		btnDowloadReport_1.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				cal = Calendar.getInstance();
				calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
				reportpath = logpathm;
				new File(reportpath + File.separator + messageboxtitle + " report").mkdirs();

				File file = new File(reportpath + File.separator + messageboxtitle + " report" + File.separator
						+ calendertime + "report.csv");

				try {
					FileWriter outputfile = new FileWriter(file);

					CSVWriter writer = new CSVWriter(outputfile);

					String[] header = { "From", "To", "File Name", "Status", "Duration", "Message Count", "Path" };

					writer.writeNext(header);

					for (int i = 0; i < table_fileConvertionreport_panel4.getRowCount(); i++) {
						String[] data1 = { table_fileConvertionreport_panel4.getValueAt(i, 0).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 1).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 2).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 3).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 4).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 5).toString(),
								table_fileConvertionreport_panel4.getValueAt(i, 6).toString() };
						writer.writeNext(data1);
					}

					writer.close();
					file.setReadOnly();
					Desktop desktop = Desktop.getDesktop();
					desktop.open(file);

				} catch (Exception e) {

				}
			}
		});
		btnDowloadReport_1.setFont(new Font("Tahoma", Font.BOLD, 15));
		btnDowloadReport_1.setBounds(443, 458, 156, 38);
		panel_4.add(btnDowloadReport_1);

		JButton btnConvertAgain = new JButton("");
		btnConvertAgain.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
			}
		});
		btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
		btnConvertAgain.setBorderPainted(false);
		btnConvertAgain.setContentAreaFilled(false);
		btnConvertAgain.setDefaultCapable(false);
		btnConvertAgain.setFocusTraversalKeysEnabled(false);
		btnConvertAgain.setFocusable(false);
		btnConvertAgain.setOpaque(false);
		btnConvertAgain.setRolloverEnabled(false);
		btnConvertAgain.setRequestFocusEnabled(false);
		btnConvertAgain.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				filetype = "";

				path = "";

				checkconvertagain = true;

				stop = false;

				btnStop.setVisible(false);

				btn_Destination.setEnabled(true);

				btn_previous_p3.setEnabled(true);

				comboBox_fileDestination_type.setEnabled(true);

				btn_Destination.setEnabled(true);

				btn_previous_p3.setEnabled(true);

				lbl_progressreport.setText("");

				dateChooser_calender_start.setEnabled(true);

				chckbxRemoveDuplicacy.setEnabled(true);

				dateChooser_calendar_end.setEnabled(true);

				dateChooser_mail_fromdate.setEnabled(true);

				dateChooser_mail_tilldate.setEnabled(true);

				dateChooser_task_start_date.setEnabled(true);

				dateChooser_task_end_date.setEnabled(true);

				btn_Destination.setEnabled(true);

				btn_previous_p3.setEnabled(true);

				btn_signout_p3.setVisible(false);

				label_11.setVisible(false);

				chckbx_Mail_Filter.setEnabled(true);

				chckbx_calender_box.setEnabled(true);

				task_box.setEnabled(true);

				CardLayout card = (CardLayout) Cardlayout.getLayout();

				card.show(Cardlayout, "panel_3");
			}
		});
		btnConvertAgain.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnConvertAgain.setBounds(442, 556, 156, 38);
		// panel_4.add(btnConvertAgain);

		JLabel label_7 = new JLabel("");
		label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_7.setBounds(0, 545, 1075, 73);
		panel_4.add(label_7);

		JLabel label_4 = new JLabel("");
		label_4.setIcon(new ImageIcon(Main_Frame.class.getResource("/topbar.png")));
		label_4.setBounds(0, 0, 1075, 71);
		contentPane.add(label_4);
	}

	void readafile(File filearray, Map.Entry<String, List<String>> entry, DefaultTableModel yourModel_1)
			throws Exception {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				readafile(files[j], entry, yourModel_1);

			} else {

				if (entry.getKey().trim().equalsIgnoreCase(files[j].getAbsolutePath())) {

					yourModel_1.addRow(new Object[] { new File(entry.getKey()).getName(), username_p3 });

				}

			}

		}

	}

	void filter_file() throws Exception {

		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
		files = null;

		jFileChooser.setMultiSelectionEnabled(true);
		jFileChooser.setAcceptAllFileFilterUsed(false);
		// jFileChooser.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter;
		if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

			filter = new FileNameExtensionFilter(".ost", "ost");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {

			filter = new FileNameExtensionFilter(".pst", "pst");

			jFileChooser.addChoosableFileFilter(filter);

			jFileChooser.addChoosableFileFilter(new FileNameExtensionFilter(".ost", "ost"));

		} else if (fileoptionm.equalsIgnoreCase("EML Files (.eml)")) {

			filter = new FileNameExtensionFilter(".eml", "eml");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")) {
			filter = new FileNameExtensionFilter(".emlx", "emlx");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {
			filter = new FileNameExtensionFilter(".msg", "msg");

			jFileChooser.addChoosableFileFilter(filter);

		} else if (fileoptionm.equalsIgnoreCase("MBOX")) {

			jFileChooser.setFileFilter(new FileNameExtensionFilter(".mbox", "mbx", "mbox"));
			jFileChooser.setAcceptAllFileFilterUsed(true);

		} else if (fileoptionm.equalsIgnoreCase("OLM files (.olm)")) {
			filter = new FileNameExtensionFilter(".olm", "olm");

			jFileChooser.addChoosableFileFilter(filter);

		}

		if (jFileChooser.showOpenDialog(restore.this) == JFileChooser.APPROVE_OPTION) {

			files = jFileChooser.getSelectedFiles();

			for (int i = 0; i < files.length; i++) {
				String extension = getFileExtension(files[i]);
				if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
						|| fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					try {

						if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
							if (extension.equalsIgnoreCase("pst") || extension.equalsIgnoreCase("ost")) {

								hashset.add(files[i]);

							}
						} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

							if (extension.equalsIgnoreCase("ost")) {
								hashset.add(files[i]);
							}
						}

					} catch (Exception e1) {

						e1.printStackTrace();
					}

				} else if (fileoptionm.equalsIgnoreCase("MBOX")) {
					hashset.add(files[i]);
				} else if (fileoptionm.equalsIgnoreCase("OLM files (.olm)")) {

					hashset.add(files[i]);
				} else {
					if (fileoptionm.equalsIgnoreCase("EML Files (.eml)")) {
						if (extension.equals("eml")) {

							hashset.add(files[i]);

						}
					} else if (fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")) {
						if (extension.equalsIgnoreCase("emlx")) {

							hashset.add(files[i]);
						}
					} else if (fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {
						if (extension.equalsIgnoreCase("msg")) {

							hashset.add(files[i]);
						}
					}

				}

			}
			DefaultTableModel model = (DefaultTableModel) table.getModel();

			while (model.getRowCount() > 0) {

				for (int i = 0; i < model.getRowCount(); ++i) {

					model.removeRow(i);
					filesno--;
				}
			}

			Iterator<File> itr = hashset.iterator();
			while (itr.hasNext()) {

				modeli = (DefaultTableModel) table.getModel();
				File fo = itr.next();
				String filet = "";
				if (fo.isFile()) {
					filet = "File";
				} else {
					filet = "Folder";
				}
				long sizeInBytes = fo.length();
				modeli.addRow(new Object[] { "<html><b>" + filesno, "<html><b>" + fo.getName(),
						"<html><b>" + fo.getAbsolutePath(), "<html><b>" + filet,
						"<html><b>" + bytes2String(sizeInBytes) });
				filesno++;
				countforfile++;
			}
		}

	}

	void destinationPath() throws Exception {
		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

		jFileChooser.setMultiSelectionEnabled(true);

		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		jFileChooser.showOpenDialog(restore.this);
		checkdestination = false;
		File file = jFileChooser.getSelectedFile();

		String destination = file.getAbsolutePath();

		tf_Destination_Location.setText(destination);

	}

	private static void expandAllNodes() {
		int j = tree.getRowCount();
		int i = 0;
		while (i < j) {
			tree.expandRow(i);
			i += 1;

		}
	}

	public void readAnOST_PstFile() {

		try {
			pst = PersonalStorage.fromFile(filepath);

			FolderInfoCollection folderInfoCollection = pst.getRootFolder().getSubFolders();

			DefaultMutableTreeNode n = new DefaultMutableTreeNode("<html><b>" + filepath);

			mainnode.add(n);

			FolderInfo folderInfo1 = pst.getRootFolder();
			String rootname = folderInfo1.getDisplayName().replaceAll("[\\[\\]]", "");
			if (rootname.equalsIgnoreCase("")) {
				rootname = "Top of Personal Folders";
			}

			DefaultMutableTreeNode node1 = new DefaultMutableTreeNode("<html><b>" + rootname);
			n.add(node1);

			for (int i = 0; i < folderInfoCollection.size(); i++) {
				if (mf.stop_tree) {
					break;
				}
				FolderInfo folderInfo = folderInfoCollection.get_Item(i);

				String foldername = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

				DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + foldername);

				node1.add(node);
				obTh.ob.MessageLabel.setText(foldername);
				if (folderInfo.hasSubFolders()) {

					readOstpstsubfolder(folderInfo, node);

				}

			}
		} catch (Exception e) {
			if (e.getMessage().contains("File not found File")) {
				JOptionPane.showMessageDialog(mf, "File is in use ", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}

			else {
				JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}
		}
	}

	public void readOstpstsubfolder(FolderInfo f, DefaultMutableTreeNode node) {

		FolderInfoCollection folderCollection = f.getSubFolders();

		for (int i = 0; i < folderCollection.size(); i++) {

			if (mf.stop_tree) {
				break;
			}
			FolderInfo folderInfo = folderCollection.get_Item(i);
			String foldername = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
			DefaultMutableTreeNode nod1 = new DefaultMutableTreeNode("<html><b>" + foldername);

			node.add(nod1);

			obTh.ob.MessageLabel.setText(foldername);
			if (folderInfo.hasSubFolders()) {

				readOstpstsubfolder(folderInfo, nod1);

			}

		}
	}

	public void readolmFile() {
		OlmStorage storage = null;

		try {
			storage = new OlmStorage(filepath);

			DefaultMutableTreeNode e = new DefaultMutableTreeNode("<html><b>" + filepath);
			mainnode.add(e);

			try {
				for (OlmFolder folder : storage.getFolderHierarchy()) {

					if (mf.stop_tree) {
						break;
					}
					String foldername = folder.getName().replaceAll("[\\[\\]]", "");
					DefaultMutableTreeNode c = new DefaultMutableTreeNode("<html><b>" + foldername);

					e.add(c);
					obTh.ob.MessageLabel.setText(foldername);
					if (folder.getSubFolders().size() > 0) {

						getFolder(folder, c);

					}

				}

			} catch (Exception e1) {

				return;
			} finally {
				storage.dispose();

			}
		} catch (Exception e) {
			if (e.getMessage().contains("File not found File")) {
				JOptionPane.showMessageDialog(mf, "File is in use ", messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}

			else {
				JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}
		}

	}

	private void getFolder(OlmFolder folder, DefaultMutableTreeNode node1) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			if (mf.stop_tree) {
				break;
			}
			String foldername = subFolder.getName().replaceAll("[\\[\\]]", "");
			DefaultMutableTreeNode nd = new DefaultMutableTreeNode("<html><b>" + foldername);
			node1.add(nd);
			obTh.ob.MessageLabel.setText(foldername);
			if (subFolder.getSubFolders().size() > 0) {
				getFolder(subFolder, nd);
			}

		}
	}

	public void readMboxFile() {

		file = new File(filepath);

		String filepath = filepath(file);

		visitAllNodes(mainnode);
		String filenamembox = file.getName();

		if (listst.contains(filepath)) {
			DefaultMutableTreeNode nd = null;
			for (int k = 0; k < lists.size(); k++) {

				if (listst.get(k).equalsIgnoreCase(filepath)) {
					nd = lists.get(k);
					break;
				}

			}

			DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + filenamembox);

			nd.add(child);

		} else {

			DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + filepath);

			mainnode.add(node);

			DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + filenamembox);

			node.add(child);
		}
		// obTh.ob.MessageLabel.setText(filenamembox);
	}

	public void readmailFile() {

		file = new File(filepath);

		String filepath = filepath(file);

		visitAllNodes(mainnode);
		String filenamemail = file.getName().replaceAll("[\\[\\]]", "");
		if (listst.contains(filepath)) {
			DefaultMutableTreeNode nd = null;
			for (int k = 0; k < lists.size(); k++) {

				if (listst.get(k).equalsIgnoreCase(filepath)) {
					nd = lists.get(k);
					break;
				}

			}

			DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + filenamemail);

			nd.add(child);

		} else {

			DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + filepath);

			mainnode.add(node);

			DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + filenamemail);

			node.add(child);
		}
		obTh.ob.MessageLabel.setText(filenamemail);
	}

	public void fileInformation_on_mail() {
		MailMessage message = null;

		try {
			MailMessage message1 = MailMessage.load(path2);

			MailConversionOptions option = new MailConversionOptions();
			MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
			message = msg.toMailMessage(option);

			String from = "";
			String Subject = "";
			String Date = "";

			try {
				from = message.getFrom().toString();
			} catch (Exception e) {

			}
			try {
				Subject = message.getSubject();
			} catch (Exception e) {

			}
			try {
				Date = message.getDate().toString();
			} catch (Exception e) {

			}

			if (message.getAttachments().size() > 0) {
				ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
				JLabel imagelabl = new JLabel();
				imagelabl.setIcon(icon);
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
			} else {
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
			}
		} catch (Exception e) {
			JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  " + filepath,
					messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));

		}
	}

	public void fileInhformation_Ost_Pst() throws Exception {
		FolderInfo f1 = pst.getRootFolder();
		String f1nmae = f1.getDisplayName().replaceAll("[\\[\\]]", "");
		if (foldername.equals(f1nmae)) {
			MessageInfoCollection messageInfoCollection = f1.getContents();
			int i2 = 0;
			for (int j = 0; j < messageInfoCollection.size(); j++)

			{

				try {
					if (Stoppreview) {
						break;
					}

					MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);
					MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
					MailConversionOptions de = new MailConversionOptions();
					MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
					MailMessage mess = contact1.toMailMessage(de);
					MapiMessage contact = MapiMessage.fromMailMessage(mess, d);
					lblTotalMessageCount.setText("Total Message Count : " + i2);
					i2++;
					listmapi.add(contact);
					String from = "";
					String Subject = "";
					Date DeliveryTime = null;
					try {
						from = contact.getSenderEmailAddress();
					} catch (Exception a) {
						from = "";
					}
					try {
						Subject = contact.getSubject();
					} catch (Exception a) {
						Subject = "";
					}
					try {
						DeliveryTime = contact.getDeliveryTime();
					} catch (Exception a) {

					}
					if (contact.getAttachments().size() > 0) {
						ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
						JLabel imagelabl = new JLabel();
						imagelabl.setIcon(icon);
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
								"<html><b>" + DeliveryTime, imagelabl });
					} else {
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(
								new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + DeliveryTime });
					}

				} catch (Exception e) {
					continue;
				}

			}

			path = "";

		}

		FolderInfoCollection folderInfoCollection = pst.getRootFolder().getSubFolders();

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {
				if (Stoppreview) {
					break;
				}
				FolderInfo f = folderInfoCollection.get_Item(i);
				path = f1nmae + File.separator + f.getDisplayName().replaceAll("[\\[\\]]", "");
				int size = f.getContentCount();

				foldername = foldername.replace("[" + size + "]", "");

				if (foldername.equals(path)) {
					folderInfo = folderInfoCollection.get_Item(i);
					MessageInfoCollection messageInfoCollection = f.getContents();
					int i2 = 0;
					for (int j = 0; j < messageInfoCollection.size(); j++)

					{
						try {

							if (Stoppreview) {
								break;
							}
							MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);

							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
							MailMessage mess = contact1.toMailMessage(de);
							MapiMessage contact = MapiMessage.fromMailMessage(mess, d);
							lblTotalMessageCount.setText("Total Message Count : " + i2);
							i2++;
							listmapi.add(contact);
							listPSTOSTgemesingo.add(messageInfo);

							String from = messageInfo.getSenderRepresentativeName();

							Date DeliveryTime = contact.getDeliveryTime();

							String Subject = messageInfo.getSubject();

							if (contact.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}
						} catch (Exception e) {
							continue;
						}

					}

					path = "";
					break;

				}
				if (f.hasSubFolders()) {
					fileInhformationsubfolder_Ost_Pst(f);

				}
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void fileInhformationsubfolder_Ost_Pst(FolderInfo folder) {
		FolderInfoCollection folderInfoCollection = folder.getSubFolders();

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {
				if (Stoppreview) {
					break;
				}
				FolderInfo folderInf = folderInfoCollection.get_Item(i);
				int size = folderInf.getContentCount();
				foldername = foldername.replace("[" + size + "]", "");
				path = path + File.separator + folderInf.getDisplayName().replaceAll("[\\[\\]]", "");
				// System.out.println(path);
				if (foldername.equals(path)) {
					folderInfo = folderInfoCollection.get_Item(i);
					MessageInfoCollection messageInfoCollection = folderInf.getContents();
					int i2 = 0;
					for (int j = 0; j < messageInfoCollection.size(); j++)

					{

						try {
							if (Stoppreview) {
								break;
							}

							MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);

							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
							MailMessage mess = contact1.toMailMessage(de);
							MapiMessage contact = MapiMessage.fromMailMessage(mess, d);

							listmapi.add(contact);

							listPSTOSTgemesingo.add(messageInfo);
							lblTotalMessageCount.setText("Total Message Count : " + i2);
							i2++;
							String from = messageInfo.getSenderRepresentativeName();

							Date DeliveryTime = contact.getDeliveryTime();

							String Subject = messageInfo.getSubject();

							if (contact.getAttachments().size() > 0) {

								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));

								JLabel imagelabl = new JLabel();

								imagelabl.setIcon(icon);

								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}
						} catch (Exception e) {
							continue;
						}

					}

					path = "";
					break;

				}
				if (folderInf.hasSubFolders()) {
					fileInhformationsubfolder_Ost_Pst(folderInf);

				}
				path = path.replace(File.separator + folderInf.getDisplayName(), "");
			} catch (Exception e) {
				continue;
			}
		}

	}

	public void fileInformation_on_mbox() {
		MboxrdStorageReader mbox = null;

		try {
			FileStream stream = new FileStream(path2, FileMode.OpenOrCreate, FileAccess.Read);

			mbox = new MboxrdStorageReader(stream.toInputStream(), false);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(mf, "File is Currupted  Please Choose another file  " + filepath,
					messageboxtitle, JOptionPane.INFORMATION_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));

		}

		MailMessage message1 = mbox.readNextMessage();
		int i2 = 0;
		while (message1 != null) {

			if (Stoppreview) {
				break;
			}
			MailConversionOptions option = new MailConversionOptions();
			MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
			MailMessage message = msg.toMailMessage(option);

			String from = message.getFrom().toString();

			String Subject = message.getSubject();

			String Date = message.getDate().toString();
			lblTotalMessageCount.setText("Total Message Count : " + i2);
			i2++;
			listmail.add(message);
			if (message.getAttachments().size() > 0) {
				ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
				JLabel imagelabl = new JLabel();
				imagelabl.setIcon(icon);
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
			} else {
				mode = (DefaultTableModel) table_fileinformation.getModel();

				mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
			}

			try {
				message1 = mbox.readNextMessage();

			} catch (Exception e) {
				continue;
			}

		}

	}

	public void fileinformation_olm() {
		OlmStorage storage = new OlmStorage(path2);
		// System.out.println("hello");

		for (OlmFolder folder : storage.getFolderHierarchy()) {

			String pa1 = folder.getName();

			if (folder.getName().equals(foldername)) {

				Iterator<MapiMessage> it = storage.enumerateMessages(folder).iterator();
				if (Stoppreview) {
					break;
				}
				if (folder.hasMessages()) {
					while (it.hasNext()) {
						MapiMessage msg1 = it.next();
						MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
						MailConversionOptions de = new MailConversionOptions();

						MailMessage mess = msg1.toMailMessage(de);
						MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

						listmapi.add(msg);
						String from = msg.getSenderEmailAddress();

						Date DeliveryTime = msg.getDeliveryTime();

						String Subject = msg.getSubject();

						if (msg.getAttachments().size() > 0) {
							ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
							JLabel imagelabl = new JLabel();
							imagelabl.setIcon(icon);
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime, imagelabl });
						} else {
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime });
						}

					}
				}
			}

			if (folder.getSubFolders().size() > 0) {

				getFolderolminfo(folder, pa1, storage);

			}

		}

	}

	private void getFolderolminfo(OlmFolder folder, String rootFolder, OlmStorage storage) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			String curpath = rootFolder + File.separator + subFolder.getName();

			// System.out.println(curpath);

			if (curpath.equals(foldername)) {

				if (subFolder.hasMessages()) {
					for (MapiMessage msg1 : storage.enumerateMessages(subFolder)) {
						try {
							if (Stoppreview) {
								break;
							}

							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();

							MailMessage mess = msg1.toMailMessage(de);
							MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

							listmapi.add(msg);
							String from = msg.getSenderEmailAddress();

							Date DeliveryTime = msg.getDeliveryTime();

							String Subject = msg.getSubject();

							if (msg.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}
						} catch (Exception e) {
							mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
							continue;
						}
					}
				}

				break;

			}

			if (subFolder.getSubFolders().size() > 0) {

				getFolderolminfo(subFolder, curpath, storage);

			}

			curpath = curpath.replace(File.separator + subFolder.getName(), "");
		}

	}

	public static void getTree(String folderName) {

		DefaultMutableTreeNode node = null;

		if (folderName.contains(File.separator)) {

			String parts[] = folderName.split(Matcher.quoteReplacement(File.separator));

			DefaultMutableTreeNode par = new DefaultMutableTreeNode(parts[parts.length - 2]);
			{
				DefaultMutableTreeNode child = new DefaultMutableTreeNode(parts[parts.length - 1]);

				search(root, par);

				lastNode.add(child);

				expandAllNodes();
			}
		}

		else {
			node = new DefaultMutableTreeNode(folderName);
			model.insertNodeInto(node, root, root.getChildCount());

			lastNode = node;
			expandAllNodes();
		}

	}

	void read_mbox_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_mbox_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);

				if (extension.equalsIgnoreCase("mbox") || extension.equalsIgnoreCase("mbx") || extension.equals("")) {
					filepath = files[j].getAbsolutePath();

					readMboxFile();

				}

			}

		}

	}

	void reademl_emlx_msg_folder(File filearray) {
		File[] files = filearray.listFiles();
		int messagesize = files.length;

		for (int j = 0; j < messagesize; j++) {

			if (files[j].isDirectory()) {

				reademl_emlx_msg_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);
				if (fileoptionm.equalsIgnoreCase("EML Files (.eml)")) {
					if (extension.equals("eml")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("EMLX Files (.emlx)")) {
					if (extension.equalsIgnoreCase("emlx")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();

					}
				} else if (fileoptionm.equalsIgnoreCase("Message Files (.msg)")) {
					if (extension.equalsIgnoreCase("msg")) {

						filepath = files[j].getAbsolutePath();

						readmailFile();
					}
				}

			}

		}
		// System.out.println("Process Completed");

	}

	void read_olm_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_olm_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);

				if (extension.equalsIgnoreCase("olm")) {
					filepath = files[j].getAbsolutePath();

					readolmFile();

				}

			}

		}

	}

	void read_PSTOST_folder(File filearray) {
		File[] files = filearray.listFiles();

		for (int j = 0; j < files.length; j++) {

			if (files[j].isDirectory()) {

				read_PSTOST_folder(files[j]);

			} else {

				String extension = getFileExtension(files[j]);
				if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
					if (extension.equalsIgnoreCase("pst") || extension.equalsIgnoreCase("ost")) {
						filepath = files[j].getAbsolutePath();
						// System.out.println(filepath);
						readAnOST_PstFile();

					}

				} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					if (extension.equalsIgnoreCase("ost")) {
						filepath = files[j].getAbsolutePath();

						readAnOST_PstFile();

					}
				}
			}

		}

	}

	public void ConvertPSTOST_365(String folde) {
		pst = PersonalStorage.fromFile(filepath);
		FolderInfo folderInfo2 = pst.getRootFolder();
		Folder = folderInfo2.getDisplayName().replaceAll("[\\[\\]]", "");

		path = Folder;
		path1 = Folder;

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
		Folderuri = clientforexchange_output.createFolder(Folderuri, path).getUri();
		parentfolder = Folderuri;
		int messagesize1;
		listduplicacy.clear();
		if (demo) {
			if (messageInfoCollection1.size() <= 50) {
				messagesize1 = messageInfoCollection1.size();
			} else {
				messagesize1 = 50;
			}

		} else {
			messagesize1 = messageInfoCollection1.size();
		}
		System.out.println("message size : " + messagesize1);
		String mailfolder = "";
		for (int i = 0; i < messagesize1; i++) {

			if (stop) {
				break;
			}
			if ((i % 100) == 0) {
				System.gc();
			}
			try {
				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);

				Date reciveddate = message.getDeliveryTime();
				if (message.getMessageClass().equals("IPM.Contact")) {

					try {
						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getContactsUri(), calendertime + "/" + Folder,
								subfolderInfo)) {
							mailfolder = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
											calendertime + "/" + Folder, null, "IPF.Contact")
									.getUri();
						}

						MapiContact con = (MapiContact) message.toMapiMessageItem();

						clientforexchange_output.createContact(mailfolder, con);
						count_destination++;

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
						if (e.getMessage()
								.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
										"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
							Progressbar.setVisible(false);
							connectionHandle(e.getMessage());
							i--;
						}
						mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " "
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

					try {

						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getCalendarUri(), calendertime + "/" + Folder,
								subfolderInfo)) {
							mailfolder = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
											calendertime + "/" + Folder, null, "IPF.Appointment")
									.getUri();
						}

						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
						cal.save(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
								AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics");
						AppointmentLoadOptions optiona = new AppointmentLoadOptions();
						optiona.getIgnoreSmtpAddressCheck();
						Appointment calDoc = Appointment
								.load(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics", optiona);

						if (chckbx_calender_box.isSelected()) {
							if (reciveddate.after(Calenderfilterstartdate)
									&& reciveddate.before(Calenderfilterenddate)) {
								clientforexchange_output.createAppointment(calDoc, mailfolder);
								count_destination++;
							}

						} else {
							clientforexchange_output.createAppointment(calDoc, mailfolder);
							count_destination++;
						}

						file.delete();
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
						if (e.getMessage()
								.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
										"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
							Progressbar.setVisible(false);
							connectionHandle(e.getMessage());
							i--;
						}
						mf.logger.warning("Exception :" + e.getMessage() + "Calendar" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.StickyNote")
						|| message.getMessageClass().equals("IPM.Task")) {
					try {
						MapiTask task = (MapiTask) message.toMapiMessageItem();
						MailConversionOptions options = new MailConversionOptions();
						options.setConvertAsTnef(true);
						String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();

						reciveddate = task.getStartDate();
						if (task_box.isSelected()) {
							if (reciveddate.after(taskfilterstartdate) && reciveddate.before(taskfilterenddate))
								clientforexchange_output.createTask(taskuri, task);
							count_destination++;
						} else {
							clientforexchange_output.createTask(taskuri, task);
							count_destination++;
						}

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}

				} else {
					try {
						String messageid = mailexchange(mess1, reciveddate, clientforexchange_output, parentfolder,
								folde);

						if (!folde.equalsIgnoreCase("Archive Folder")) {
							if (((message.getFlags()
									& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
								clientforexchange_output.setReadFlag(messageid, true);

							} else {
								clientforexchange_output.setReadFlag(messageid, false);
							}
						}

					}

					catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
						if (e.getMessage().equalsIgnoreCase(
								"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
								|| e.getMessage().contains(
										"No connection could be made because the target machine actively refused it.")
								|| e.getMessage().contains("ConnectFailure")
								|| e.getMessage().contains("Rate limit hit")
								|| e.getMessage().contains("Operation failed")
								|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
										"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
							Progressbar.setVisible(false);
							connectionHandle(e.getMessage());
							i--;
						}
						mf.logger.warning("Exception :" + e.getMessage() + "Message " + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						continue;
					}

				}

				lbl_progressreport.setText("  Total Message Saved Count  " + count_destination + "  " + Folder
						+ "   Extarcting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {

				if (stop) {
					break;
				}
				String subfolder = Folderuri;
				// int count = 1;
				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

				path = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(path)) {

						lbl_progressreport.setText(" Getting Folder " + Folder);
						// progressBar_message_p3.setStringPainted(true);
						// progressBar_message_p3.setValue(0);

						subfolder = clientforexchange_output.createFolder(subfolder, Folder).getUri();

						MessageInfoCollection messageInfoCollection = folderInfo.getContents();
						listduplicacy.clear();
						int messagesize;
						if (demo) {
							if (messageInfoCollection.size() <= 50) {
								messagesize = messageInfoCollection.size();
							} else {
								messagesize = 50;
							}

						} else {
							messagesize = messageInfoCollection.size();
						}

						for (int i = 0; i < messagesize; i++) {
							if (stop) {
								break;
							}
							if ((i % 100) == 0) {
								System.gc();
							}
							try {
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								Date reciveddate = message.getDeliveryTime();
								if (message.getMessageClass().equals("IPM.Contact")) {

									try {
										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getContactsUri(),
												calendertime + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output.createFolder(
													clientforexchange_output.getMailboxInfo().getContactsUri(),
													calendertime + "/" + Folder, null, "IPF.Contact").getUri();
										}

										MapiContact con = (MapiContact) message.toMapiMessageItem();

										clientforexchange_output.createContact(mailfolder, con);
										count_destination++;

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().contains(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											// System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " "
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

									try {

										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getCalendarUri(),
												calendertime + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output
													.createFolder(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															calendertime + "/" + Folder, null, "IPF.Appointment")
													.getUri();
										}

										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
										cal.save(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
												AppointmentSaveFormat.Ics);
										File file = new File(
												temppathm + File.separator + mf.namingconventionmapi(message) + ".ics");
										AppointmentLoadOptions optiona = new AppointmentLoadOptions();
										optiona.getIgnoreSmtpAddressCheck();
										Appointment calDoc = Appointment.load(
												temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
												optiona);

										if (chckbx_calender_box.isSelected()) {
											if (reciveddate.after(Calenderfilterstartdate)
													&& reciveddate.before(Calenderfilterenddate)) {
												clientforexchange_output.createAppointment(calDoc, mailfolder);
												count_destination++;
											}

										} else {
											clientforexchange_output.createAppointment(calDoc, mailfolder);
											count_destination++;
										}

										file.delete();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().contains(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											// System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Calendar" + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.StickyNote")
										|| message.getMessageClass().equals("IPM.Task")) {
									try {
										MapiTask task = (MapiTask) message.toMapiMessageItem();
										MailConversionOptions options = new MailConversionOptions();
										options.setConvertAsTnef(true);
										String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();

										reciveddate = task.getStartDate();
										if (task_box.isSelected()) {
											if (reciveddate.after(mailfilterstartdate)
													&& reciveddate.before(mailfilterenddate))
												clientforexchange_output.createTask(taskuri, task);
											count_destination++;
										} else {
											clientforexchange_output.createTask(taskuri, task);
											count_destination++;
										}

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										e.printStackTrace();
										continue;
									}

								} else {
									try {
										String messageid = mailexchange(mess, reciveddate, clientforexchange_output,
												subfolder, folde);
										if (!folde.equalsIgnoreCase("Archive Folder")) {
											if (((message.getFlags()
													& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
												clientforexchange_output.setReadFlag(messageid, true);

											} else {
												clientforexchange_output.setReadFlag(messageid, false);
											}
										}

									}

									catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().equalsIgnoreCase(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Message " + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								}

								lbl_progressreport.setText("  Total Message Saved Count  " + count_destination + "  "
										+ Folder + "   Extarcting messsage " + message.getSubject());

							} catch (Exception e) {
								continue;
							}

						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderpstost_exchange(folderInfo, subfolder, folde);

				}
			} catch (Exception e) {
				continue;
			}

		}

	}

	public void getsubfolderpstost_exchange(FolderInfo folderi, String p, String folde) {
		FolderInfoCollection subfolder1 = folderi.getSubFolders();
		for (int k = 0; k < subfolder1.size(); k++) {

			try {
				if (stop) {
					break;
				}
				// int count = 1;
				FolderInfo folderInfo = subfolder1.get_Item(k);
				String Folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
				// progressBar_message_p3.setStringPainted(true);
				// progressBar_message_p3.setValue(0);

				String mailfolder = "";
				path = path + File.separator + Folder;

				String subfolder = p;
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(path)) {
						lbl_progressreport.setText("Getting : " + Folder);

						subfolder = clientforexchange_output.createFolder(p, Folder).getUri();
						listduplicacy.clear();
						// System.out.println("path for folder : " + path);

						MessageInfoCollection messageInfoCollection = folderInfo.getContents();

						int messagesize;
						if (demo) {
							if (messageInfoCollection.size() <= 50) {
								messagesize = messageInfoCollection.size();
							} else {
								messagesize = 50;
							}

						} else {
							messagesize = messageInfoCollection.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if (stop) {
									break;
								}
								if ((i % 100) == 0) {
									System.gc();
								}
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								Date reciveddate = message.getDeliveryTime();
								if (message.getMessageClass().equals("IPM.Contact")) {

									try {
										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getContactsUri(),
												calendertime + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output.createFolder(
													clientforexchange_output.getMailboxInfo().getContactsUri(),
													calendertime + "/" + Folder, null, "IPF.Contact").getUri();
										}

										MapiContact con = (MapiContact) message.toMapiMessageItem();

										String ConData = clientforexchange_output.createContact(mailfolder, con);
										clientforexchange_output.archiveItem(mailfolder, ConData);
										count_destination++;

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().contains(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											// System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " "
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

									try {

										ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
										if (!clientforexchange_output.folderExists(
												clientforexchange_output.getMailboxInfo().getCalendarUri(),
												calendertime + "/" + Folder, subfolderInfo)) {
											mailfolder = clientforexchange_output
													.createFolder(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															calendertime + "/" + Folder, null, "IPF.Appointment")
													.getUri();
										}

										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
										cal.save(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
												AppointmentSaveFormat.Ics);
										File file = new File(
												temppathm + File.separator + mf.namingconventionmapi(message) + ".ics");
										AppointmentLoadOptions optiona = new AppointmentLoadOptions();
										optiona.getIgnoreSmtpAddressCheck();
										Appointment calDoc = Appointment.load(
												temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
												optiona);

										if (chckbx_calender_box.isSelected()) {
											if (reciveddate.after(Calenderfilterstartdate)
													&& reciveddate.before(Calenderfilterenddate)) {
												clientforexchange_output.createAppointment(calDoc, mailfolder);
												count_destination++;
											}

										} else {
											clientforexchange_output.createAppointment(calDoc, mailfolder);
											count_destination++;
										}

										file.delete();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().contains(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											// System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Calendar" + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.StickyNote")
										|| message.getMessageClass().equals("IPM.Task")) {
									try {
										MapiTask task = (MapiTask) message.toMapiMessageItem();
										MailConversionOptions options = new MailConversionOptions();
										options.setConvertAsTnef(true);
										String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();

										reciveddate = task.getStartDate();
										if (task_box.isSelected()) {
											if (reciveddate.after(mailfilterstartdate)
													&& reciveddate.before(mailfilterenddate))
												clientforexchange_output.createTask(taskuri, task);
											count_destination++;
										} else {
											clientforexchange_output.createTask(taskuri, task);
											count_destination++;
										}

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										e.printStackTrace();
										continue;
									}

								} else {
									try {
										String messageid = mailexchange(mess, reciveddate, clientforexchange_output,
												subfolder, folde);
										if (!folde.equalsIgnoreCase("Archive Folder")) {
											if (((message.getFlags()
													& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
												clientforexchange_output.setReadFlag(messageid, true);

											} else {
												clientforexchange_output.setReadFlag(messageid, false);
											}
										}
									}

									catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
										if (e.getMessage().contains(
												"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
												|| e.getMessage().contains(
														"No connection could be made because the target machine actively refused it.")
												|| e.getMessage().contains("ConnectFailure")
												|| e.getMessage().contains("Operation failed")
												|| e.getMessage().contains("Rate limit hit")
												|| e.getMessage().contains("Operation has been canceled")
												|| e.getMessage().contains(
														"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
											Progressbar.setVisible(false);
											connectionHandle(e.getMessage());
											i--;
										} else {
											System.out.println(">>>>>>>>>>>> " + e.getMessage());
										}
										mf.logger.warning("Exception :" + e.getMessage() + "Message " + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										continue;
									}

								}
								lbl_progressreport.setText("  Total Message Saved Count  " + count_destination + "  "
										+ Folder + "   Extarcting messsage " + message.getSubject());

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								}
								continue;
							}

						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderpstost_exchange(folderInfo, subfolder, folde);

				}

				path = path.replace(File.separator + Folder, "");

			} catch (Exception e) {
				System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
				if (e.getMessage()
						.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
						|| e.getMessage()
								.contains("No connection could be made because the target machine actively refused it.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
								"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
					Progressbar.setVisible(false);
					connectionHandle(e.getMessage());

				} else {
					System.out.println(">>>>>>>>>>>> " + e.getMessage());
				}
				continue;
			}

		}

	}

	void ConvertMboxfolder_exchange(File filearray) throws Exception {
		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {
			if (files[i].isFile()) {

				for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
					if (stop) {
						break;
					}
					if (entry.getKey().trim().equalsIgnoreCase(files[i].getAbsolutePath())) {

						String s = Folderuri;
						foldername3 = files[i].getName().replace(".mbox", "").replace(".mbx", "");
						filepath = files[i].getAbsolutePath();

						Folderuri = clientforexchange_output.createFolder(Folderuri, foldername3).getUri();
						String folde = "";
						ConvertMbox_365(folde);
						Folderuri = s;
					}
				}

			} else {

				String s = Folderuri;
				foldername4 = files[i].getName();
				Folderuri = clientforexchange_output.createFolder(Folderuri, foldername4).getUri();
				ConvertMboxfolder_exchange(files[i]);
				Folderuri = s;

			}

		}

	}

	void ConvertPSTOST_folder_exchange(File filearray) throws Exception {
		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {
			if (files[i].isFile()) {
				calendertime = Calendar.getInstance().getTime().toString();
				String extension = getFileExtension(files[i]);

				if (fileoptionm.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {

					if (extension.equalsIgnoreCase("pst")||extension.equalsIgnoreCase("ost")) {
						for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
							if (stop) {
								break;
							}
							if (entry.getKey().trim().equalsIgnoreCase(files[i].getAbsolutePath())) {

								pstfolderlist = new ArrayList<String>();

								pstfolderlist.addAll(entry.getValue());

								String s = Folderuri;
								foldername3 = files[i].getName().replace(".pst", "").replace(".ost", "");
								filepath = files[i].getAbsolutePath();

								String user = "";
								String folde = "";

								for (int j = 0; j < table_2.getRowCount(); j++) {

									if (table_2.getValueAt(j, 0).toString().equals(file.getName())) {
										user = table_2.getValueAt(j, 1).toString();
										folde = table_2.getValueAt(j, 2).toString();
										break;
									}

								}
								fname = files[i].getName().replace(".pst", "").replace(".ost", "");
								Boolean flagcheck = true;

								try {
									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);
									clientforexchange_output.getMailboxInfo(user);
								} catch (Exception e1) {

									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);

									clientforexchange_output.impersonateUser(3, user);
								}

								String Publicfolderuri = "";
								if (folde.equalsIgnoreCase("Public Folder")) {
									ExchangeFolderInfo folderj = clientforexchange_output
											.getFolderInfo("PublicFoldersRoot");

									Publicfolderuri = folderj.getUri();
								}

								if (chckbxCustomFolderName.isSelected()) {
									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
									try {
										if (folde.equalsIgnoreCase("Archive Folder")) {

											Folderuri = clientforexchange_output.createFolder(Folderuri, customerfolder)
													.getUri();
											fa = Folderuri;
										} else if (folde.equalsIgnoreCase("Public Folder")) {
											try {
												Folderuri = clientforexchange_output
														.createFolder(Publicfolderuri, customerfolder).getUri();
												fa = Folderuri;
											} catch (Exception e) {
												JOptionPane.showMessageDialog(restore.this,
														"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder is created to perform this operation  on this object. See the folder contact or your system administrator",
														messageboxtitle, JOptionPane.ERROR_MESSAGE,
														new ImageIcon(backup.class.getResource("/information.png")));
											}
										}

									} catch (Exception e) {

										Folderuri = clientforexchange_output
												.createFolder(customerfolder + "(" + calendertime + ")").getUri();
										fa = Folderuri;

									}
								} else {

									if (folde.equalsIgnoreCase("Public Folder")) {
										try {
											Folderuri = clientforexchange_output
													.createFolder(Publicfolderuri, calendertime).getUri();
											fa = Folderuri;
										} catch (Exception e) {
											JOptionPane.showMessageDialog(restore.this,
													"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder  is created to perform this operation  on this object. See the folder contact or your system administrator",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(backup.class.getResource("/information.png")));
										}
									} else {
										if (flagcheck) {
											Folderuri = clientforexchange_output.createFolder(calendertime).getUri();
											flagcheck = false;
											fa = Folderuri;
										}
									}
								}
								if (!folde.equalsIgnoreCase("Archive Folder")) {

									if (chckbxrestoretodefault.isSelected()) {

										try {

											ConvertPST_defaultrestore();

										} catch (Exception e) {

										}

									} else {
										Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();
										try {

											ConvertPSTOST_365(folde);

										} catch (Exception e) {

										}

									}
								} else {
									if (chckbxrestoretodefault.isSelected()) {
										JOptionPane.showMessageDialog(mf,
												"As you have selected Archieve Folder as well as selected Default folder it will be stored in Archieve Folder",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
									Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();
									try {

										ConvertPSTOST_365(folde);
									} catch (Exception e) {

									}

								}

								Folderuri = s;
							}
						}
					}
				} else if (fileoptionm.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

					if (extension.equals("ost")) {
						for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
							if (stop) {
								break;
							}
							if (entry.getKey().trim().equalsIgnoreCase(files[i].getAbsolutePath())) {

								pstfolderlist = new ArrayList<String>();

								pstfolderlist.addAll(entry.getValue());
								String s = Folderuri;
								foldername3 = files[i].getName().replace(".ost", "");
								filepath = files[i].getAbsolutePath();
								String user = "";
								String folde = "";

								for (int j = 0; j < table_2.getRowCount(); j++) {

									if (table_2.getValueAt(j, 0).toString().equals(file.getName())) {
										user = table_2.getValueAt(j, 1).toString();
										folde = table_2.getValueAt(j, 2).toString();
										break;
									}

								}
								fname = files[i].getName().replace(".pst", "");
								Boolean flagcheck = true;
								calendertime = Calendar.getInstance().getTime().toString();
								try {
									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);
									clientforexchange_output.getMailboxInfo(user);
								} catch (Exception e1) {

									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);

									clientforexchange_output.impersonateUser(3, user);
								}

								String Publicfolderuri = "";
								if (folde.equalsIgnoreCase("Public Folder")) {
									ExchangeFolderInfo folderj = clientforexchange_output
											.getFolderInfo("PublicFoldersRoot");

									Publicfolderuri = folderj.getUri();
								}

								if (chckbxCustomFolderName.isSelected()) {
									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
									try {
										if (folde.equalsIgnoreCase("Archive Folder")) {

											Folderuri = clientforexchange_output.createFolder(customerfolder).getUri();
											fa = Folderuri;
										} else if (folde.equalsIgnoreCase("Public Folder")) {
											try {
												Folderuri = clientforexchange_output
														.createFolder(Publicfolderuri, customerfolder).getUri();
												fa = Folderuri;
											} catch (Exception e) {
												JOptionPane.showMessageDialog(restore.this,
														"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder is created to perform this operation  on this object. See the folder contact or your system administrator",
														messageboxtitle, JOptionPane.ERROR_MESSAGE,
														new ImageIcon(backup.class.getResource("/information.png")));
											}
										}

									} catch (Exception e) {

										Folderuri = clientforexchange_output
												.createFolder(customerfolder + "(" + calendertime + ")").getUri();
										fa = Folderuri;

									}
								} else {

									if (folde.equalsIgnoreCase("Public Folder")) {
										try {
											Folderuri = clientforexchange_output
													.createFolder(Publicfolderuri, calendertime).getUri();
											fa = Folderuri;
										} catch (Exception e) {
											JOptionPane.showMessageDialog(restore.this,
													"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder  is created to perform this operation  on this object. See the folder contact or your system administrator",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(backup.class.getResource("/information.png")));
										}
									} else {
										if (flagcheck) {
											Folderuri = clientforexchange_output.createFolder(calendertime).getUri();
											flagcheck = false;
											fa = Folderuri;
										}
									}
								}
								if (!folde.equalsIgnoreCase("Archive Folder")) {

									if (chckbxrestoretodefault.isSelected()) {
										try {

											ConvertPST_defaultrestore();
										} catch (Exception e) {

										}

									} else {
										Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();
										try {

											ConvertPSTOST_365(folde);
										} catch (Exception e) {

										}

									}
								} else {
									if (chckbxrestoretodefault.isSelected()) {
										JOptionPane.showMessageDialog(mf,
												"As you have selected Archieve Folder as well as selected Default folder it will be stored in Archieve Folder",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
									Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();

									try {

										ConvertPSTOST_365(folde);
									} catch (Exception e) {

									}

								}

								Folderuri = s;

							}
						}
					}
				} else if (fileoptionm.equalsIgnoreCase("Mbox")) {

					if (extension.equals("mbox") || extension.equals("mbx") || extension.equals("")) {
						for (Map.Entry<String, List<String>> entry : hm.entrySet()) {
							if (stop) {
								break;
							}
							if (entry.getKey().trim().equalsIgnoreCase(files[i].getAbsolutePath())) {

								foldername3 = files[i].getName().replace(".mbx", "").replace(".mbox", "");
								filepath = files[i].getAbsolutePath();
								String user = "";
								String folde = "";

								for (int j = 0; j < table_2.getRowCount(); j++) {

									if (table_2.getValueAt(j, 0).toString().equals(file.getName())) {
										user = table_2.getValueAt(j, 1).toString();
										folde = table_2.getValueAt(j, 2).toString();
										break;
									}

								}
								fname = files[i].getName().replace(".mbx", "").replace(".mbox", "");
								Boolean flagcheck = true;
								calendertime = Calendar.getInstance().getTime().toString();
								try {
									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);
									clientforexchange_output.getMailboxInfo(user);
								} catch (Exception e1) {

									clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3,
											password_p3);

									clientforexchange_output.impersonateUser(3, user);
								}

								String Publicfolderuri = "";
								if (folde.equalsIgnoreCase("Public Folder")) {
									ExchangeFolderInfo folderj = clientforexchange_output
											.getFolderInfo("PublicFoldersRoot");

									Publicfolderuri = folderj.getUri();
								}

								if (chckbxCustomFolderName.isSelected()) {
									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
									try {
										if (folde.equalsIgnoreCase("Archive Folder")) {

											Folderuri = clientforexchange_output.createFolder(customerfolder).getUri();
											fa = Folderuri;
										} else if (folde.equalsIgnoreCase("Public Folder")) {
											try {
												Folderuri = clientforexchange_output
														.createFolder(Publicfolderuri, customerfolder).getUri();
												fa = Folderuri;
											} catch (Exception e) {
												JOptionPane.showMessageDialog(restore.this,
														"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder is created to perform this operation  on this object. See the folder contact or your system administrator",
														messageboxtitle, JOptionPane.ERROR_MESSAGE,
														new ImageIcon(backup.class.getResource("/information.png")));
											}
										}

									} catch (Exception e) {

										Folderuri = clientforexchange_output
												.createFolder(customerfolder + "(" + calendertime + ")").getUri();
										fa = Folderuri;

									}
								} else {

									if (folde.equalsIgnoreCase("Public Folder")) {
										try {
											Folderuri = clientforexchange_output
													.createFolder(Publicfolderuri, calendertime).getUri();
											fa = Folderuri;
										} catch (Exception e) {
											JOptionPane.showMessageDialog(restore.this,
													"Cannot Create the folder. You do not have sufficient permission  or public mailbox folder  is created to perform this operation  on this object. See the folder contact or your system administrator",
													messageboxtitle, JOptionPane.ERROR_MESSAGE,
													new ImageIcon(backup.class.getResource("/information.png")));
										}
									} else {
										if (flagcheck) {
											Folderuri = clientforexchange_output.createFolder(calendertime).getUri();
											flagcheck = false;
											fa = Folderuri;
										}
									}
								}
								if (!folde.equalsIgnoreCase("Archive Folder")) {

									if (chckbxrestoretodefault.isSelected()) {

										try {
											ConvertMbox_defaultrestore();
										} catch (Exception e) {

										}

									} else {
										Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();

										try {
											ConvertPSTOST_365(folde);
										} catch (Exception e) {

										}

									}
								} else {
									if (chckbxrestoretodefault.isSelected()) {
										JOptionPane.showMessageDialog(mf,
												"As you have selected Archieve Folder as well as selected Default folder it will be stored in Archieve Folder",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
									Folderuri = clientforexchange_output.createFolder(Folderuri, fname).getUri();
									try {

										ConvertMbox_365(folde);
									} catch (Exception e) {

									}

								}

								Folderuri = s;

							}
						}
					}
				}

			} else {

				String s = Folderuri;
				foldername4 = files[i].getName();
				Folderuri = clientforexchange_output.createFolder(Folderuri, foldername4).getUri();
				ConvertPSTOST_folder_exchange(files[i]);
				Folderuri = s;

			}

		}

	}

	private void impersonationUser() {

		Contact[] contact = clientforexchange_output.getMailboxes();
		List<String> uscheck = new ArrayList<String>();

		DefaultTableModel dtm = (DefaultTableModel) table_userid.getModel();
		for (int j = 0; j < contact.length; j++) {

			EmailAddressList contactmailAdd = contact[j].getEmailAddresses();

			for (EmailAddress emailAddress : contactmailAdd) {

				clientforexchange_output.impersonateUser(3, emailAddress.getAddress());
				try {
					ExchangeMailboxInfo mialinfo = clientforexchange_output.getMailboxInfo();
					clientforexchange_output.getMailboxInfo(mialinfo.getMailboxUri());
					String emailaddress = emailAddress.getAddress();
					if (!uscheck.contains(emailaddress)
							&& emailaddress.contains(usernameverification[usernameverification.length - 1])) {
						uscheck.add(emailaddress);
						obThmailbox.ob.MessageLabel.setText(emailaddress);
						dtm.addRow(new Object[] { true, "", emailaddress });
					}
				} catch (Exception e) {

				}

			}

			try {
				clientforexchange_output.resetImpersonation();
				clientforexchange_output.dispose();

			} catch (Exception e) {

			}

		}

	}

	public void ConvertPST_defaultrestore() throws Exception {

		String path2 = "";
		FolderInfo folderInfo2 = pst.getRootFolder();

		Folder = folderInfo2.getDisplayName().replaceAll("[\\[\\]]", "");
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Top of Personal Folders";
		}

		path = path + File.separator + Folder;
		path2 = Folder;

		// int countr = 0;

		String defaultfolder = clientforexchange_output.createFolder(fname + calendertime).getUri();
		if (folderInfo2.getContentCount() > 0) {
			String mailbox = clientforexchange_output.createFolder(defaultfolder, Folder).getUri();

			MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();
			messageaddpst(messageInfoCollection1, mailbox, Folder);
		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		// //int count = 1;
		for (int j = 0; j < folderInf.size(); j++) {
			try {

				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

				if (stop) {
					break;
				}
				path = path + File.separator + Folder;
				String path3 = Folder;
				String mailfolder = "";
				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}
					String path1 = pstfolderlist.get(l).replace(path2 + File.separator, "");

					System.out.println(path1);

					if (path1.equalsIgnoreCase(path3)) {
						lbl_progressreport.setText(" Getting Folder " + Folder);

						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();

						if (Folder.contains("Inbox")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getInboxUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();

								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Deleted Item")) {

							if (folderInfo.getContentCount() > 0) {

								mailfolder = clientforexchange_output.getMailboxInfo().getDeletedItemsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Calendar")) {

							if (folderInfo.getContentCount() > 0) {

								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}
								ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
								if (!clientforexchange_output.folderExists(
										clientforexchange_output.getMailboxInfo().getCalendarUri(),
										calendertime + "/" + Folder, subfolderInfo)) {
									mailfolder = clientforexchange_output
											.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
													calendertime + "/" + Folder, null, "IPF.Appointment")
											.getUri();
								}
								for (int i = 0; i < messagesize; i++) {
									try {
										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										Date reciveddate = message.getDeliveryTime();
										if (message.getMessageClass().equals("IPM.Appointment")
												|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {

												MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
												cal.save(temppathm + File.separator + mf.namingconventionmapi(message)
														+ ".ics", AppointmentSaveFormat.Ics);
												File file = new File(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics");
												AppointmentLoadOptions optiona = new AppointmentLoadOptions();
												optiona.getIgnoreSmtpAddressCheck();
												Appointment calDoc = Appointment.load(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics", optiona);

												if (chckbxRemoveDuplicacy.isSelected()) {

													String input = mf.duplicacymapiCal(cal);

													if (!listdupliccal.contains(input)) {
														System.out.println("Not a duplicate message");
														listdupliccal.add(input);

														if (chckbx_calender_box.isSelected()) {
															if (reciveddate.after(Calenderfilterstartdate)
																	&& reciveddate.before(Calenderfilterenddate)) {

																clientforexchange_output.createAppointment(calDoc,
																		mailfolder);

																count_destination++;
															} else if (reciveddate.equals(Calenderfilterstartdate)
																	|| reciveddate.equals(Calenderfilterenddate)) {
																clientforexchange_output.createAppointment(calDoc,
																		mailfolder);
																count_destination++;

															}

														} else {

															clientforexchange_output.createAppointment(calDoc,
																	mailfolder);

															count_destination++;

														}

													} else {

													}
												} else {

													if (chckbx_calender_box.isSelected()) {
														if (reciveddate.after(Calenderfilterstartdate)
																&& reciveddate.before(Calenderfilterenddate)) {

															clientforexchange_output.createAppointment(calDoc,
																	mailfolder);

															count_destination++;
														} else if (reciveddate.equals(Calenderfilterstartdate)
																|| reciveddate.equals(Calenderfilterenddate)) {
															clientforexchange_output.createAppointment(calDoc,
																	mailfolder);
															count_destination++;

														}

													} else {

														clientforexchange_output.createAppointment(calDoc, mailfolder);

														count_destination++;

													}

												}
												file.delete();
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "  "
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}

							}
						} else if (Folder.contains("Tasks") || Folder.contains("ToDo")) {

							if (folderInfo.getContentCount() > 0) {

								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {
										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										Date reciveddate = message.getDeliveryTime();
										if (messageInfo.getMessageClass().equals("IPM.StickyNote")
												|| messageInfo.getMessageClass().equals("IPM.Task")) {
											try {

												MapiTask task = null;
												if (messageInfo.getMessageClass().equals("IPM.Task")) {
													task = (MapiTask) message.toMapiMessageItem();
												}
												mailfolder = clientforexchange_output.getMailboxInfo().getTasksUri();
												if (chckbxRemoveDuplicacy.isSelected()) {
													String input = "";
													if (messageInfo.getMessageClass().equals("IPM.Task")) {
														input = mf.duplicacymapiTask(task);
													}
													if (!listduplictask.contains(input)) {
														System.out.println("Not a duplicate message");
														listduplictask.add(input);

														if (task_box.isSelected()) {
															if (reciveddate.after(taskfilterstartdate)
																	&& reciveddate.before(taskfilterenddate)) {
																clientforexchange_output.appendMessage(mailfolder,
																		mess);
																count_destination++;
															} else if (reciveddate.equals(taskfilterstartdate)
																	|| reciveddate.equals(taskfilterenddate)) {
																clientforexchange_output.appendMessage(mailfolder,
																		mess);
																count_destination++;

															}

														} else {
															clientforexchange_output.appendMessage(mailfolder, mess);
															count_destination++;

														}

													}
												} else {

													if (task_box.isSelected()) {
														if (reciveddate.after(taskfilterstartdate)
																&& reciveddate.before(taskfilterenddate)) {
															clientforexchange_output.appendMessage(mailfolder, mess);
															count_destination++;
														} else if (reciveddate.equals(taskfilterstartdate)
																|| reciveddate.equals(taskfilterenddate)) {
															clientforexchange_output.appendMessage(mailfolder, mess);
															count_destination++;

														}

													} else {
														clientforexchange_output.appendMessage(mailfolder, mess);
														count_destination++;

													}

												}

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error: " + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Task" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else if (Folder.contains("Contacts")) {

							if (folderInfo.getContentCount() > 0) {

								ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
								if (!clientforexchange_output.folderExists(
										clientforexchange_output.getMailboxInfo().getContactsUri(),
										calendertime + "/" + Folder, subfolderInfo)) {
									mailfolder = clientforexchange_output
											.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
													calendertime + "/" + Folder, null, "IPF.Contact")
											.getUri();
								}

								MessageInfoCollection messageInfoCollection = folderInfo.getContents();

								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {
										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										if (message.getMessageClass().equals("IPM.Contact")) {
											try {

												MapiContact con = (MapiContact) message.toMapiMessageItem();
												ByteArrayOutputStream bos = new ByteArrayOutputStream();
												con.save(bos, ContactSaveFormat.VCard);
												ByteArrayInputStream inStream = new ByteArrayInputStream(
														bos.toByteArray());
												MapiContact mapicontact = MapiContact.fromVCard(inStream);
												if (chckbxRemoveDuplicacy.isSelected()) {

													String input = mf.duplicacymapiContact(con);

													if (!listdupliccontact.contains(input)) {
														System.out.println("Not a duplicate message");
														listdupliccontact.add(input);

														clientforexchange_output.createContact(mailfolder, mapicontact);
														count_destination++;

													}
												} else {

													clientforexchange_output.createContact(mailfolder, mapicontact);
													count_destination++;

												}

												count_destination++;

											} catch (Error e) {
												mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else if (Folder.contains("Outbox")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getOutboxUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Draft")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getDraftsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Junk Email")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getJunkeMailsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Notes")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getNotesUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Journal")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getJournalUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else if (Folder.contains("Sent")) {

							if (folderInfo.getContentCount() > 0) {
								mailfolder = clientforexchange_output.getMailboxInfo().getSentItemsUri();
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();
								messageaddpst(messageInfoCollection, mailfolder, Folder);
							}
						} else {
							mailfolder = clientforexchange_output.createFolder(defaultfolder, Folder).getUri();
							if (folderInfo.getContentCount() > 0) {
								MessageInfoCollection messageInfoCollection = folderInfo.getContents();

								listduplicacy.clear();
								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}
								String scm = "";

								for (int i = 0; i < messagesize; i++) {
									try {
										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										Date reciveddate = message.getDeliveryTime();
										if (message.getMessageClass().equals("IPM.Contact")) {
											try {
												ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };

												if (!clientforexchange_output.folderExists(
														clientforexchange_output.getMailboxInfo().getContactsUri(),
														calendertime + "/" + Folder, subfolderInfo)) {
													scm = clientforexchange_output
															.createFolder(
																	clientforexchange_output.getMailboxInfo()
																			.getContactsUri(),
																	calendertime + "/" + Folder, null, "IPF.Contact")
															.getUri();
												}

												MapiContact con = (MapiContact) message.toMapiMessageItem();

												clientforexchange_output.createContact(scm, con);

												count_destination++;

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												continue;
											}

										} else if (message.getMessageClass().equals("IPM.Appointment")
												|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {

												MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

												ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
												if (!clientforexchange_output.folderExists(
														clientforexchange_output.getMailboxInfo().getCalendarUri(),
														calendertime + "/" + Folder, subfolderInfo)) {
													scm = clientforexchange_output.createFolder(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															calendertime + "/" + Folder, null, "IPF.Appointment")
															.getUri();
												}

												cal.save(temppathm + File.separator + mf.namingconventionmapi(message)
														+ ".ics", AppointmentSaveFormat.Ics);
												File file = new File(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics");
												AppointmentLoadOptions optiona = new AppointmentLoadOptions();
												optiona.getIgnoreSmtpAddressCheck();
												Appointment calDoc = Appointment.load(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics", optiona);

												if (chckbx_calender_box.isSelected()) {
													if (reciveddate.after(Calenderfilterstartdate)
															&& reciveddate.before(Calenderfilterenddate)) {
														clientforexchange_output.createAppointment(calDoc, scm);
														count_destination++;
													}

												} else {
													clientforexchange_output.createAppointment(calDoc, scm);
													count_destination++;

												}
												file.delete();
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Calendar" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										} else if (message.getMessageClass().equals("IPM.StickyNote")
												|| message.getMessageClass().equals("IPM.Task")) {
											try {
												MapiTask task = (MapiTask) message.toMapiMessageItem();
												MailConversionOptions options = new MailConversionOptions();
												options.setConvertAsTnef(true);
												String taskuri = clientforexchange_output.getMailboxInfo()
														.getTasksUri();
												if (chckbx_Mail_Filter.isSelected()) {
													if (reciveddate.after(mailfilterstartdate)
															&& reciveddate.before(mailfilterenddate))

													{
														clientforexchange_output.createTask(taskuri, task);
														count_destination++;

													}

												} else {
													clientforexchange_output.createTask(taskuri, task);
													count_destination++;
												}

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Task" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										} else {

											try {
												String messageid = mailexchange(mess, reciveddate,
														clientforexchange_output, mailfolder, "");
												if (!messageid.equalsIgnoreCase("")) {
													if (((message.getFlags()
															& MapiMessageFlags.MSGFLAG_READ) == MapiMessageFlags.MSGFLAG_READ)) {
														clientforexchange_output.setReadFlag(messageid, true);

													} else {
														clientforexchange_output.setReadFlag(messageid, false);
													}
												}

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Message" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										e.printStackTrace();
										continue;
									}

								}
							}
						}

					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOST_defaultrestore(folderInfo, path2, path3, mailfolder);
				}

			} catch (Exception e) {
				continue;
			}

		}

	}

	void psttask(MapiTask task, Date reciveddate, MapiMessage message, String taskuri, Boolean checktask,
			MessageInfo messageInfo) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = "";
			if (messageInfo.getMessageClass().equals("IPM.Task")) {
				input = mf.duplicacymapiTask(task);
			}

			if (!listduplictask.contains(input)) {

				listduplictask.add(input);

				if (checktask) {

					if (task_box.isSelected()) {

						if (reciveddate.after(taskfilterstartdate) && reciveddate.before(taskfilterenddate)) {

							clientforexchange_output.createTask(taskuri, task);

							count_destination++;

						} else if (reciveddate.equals(taskfilterstartdate) || reciveddate.equals(taskfilterenddate)) {

							clientforexchange_output.createTask(taskuri, task);

							count_destination++;

						}

					} else {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					}
				} else {
					if (task_box.isSelected()) {

						if (reciveddate.after(taskfilterstartdate) && reciveddate.before(taskfilterenddate)) {

							clientforexchange_output.createTask(taskuri, task);

							count_destination++;

						} else if (reciveddate.equals(taskfilterstartdate) || reciveddate.equals(taskfilterenddate)) {

							clientforexchange_output.createTask(taskuri, task);

							count_destination++;

						}

					} else {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					}
				}

			}
		} else {

			if (checktask) {

				if (task_box.isSelected()) {

					if (reciveddate.after(taskfilterstartdate) && reciveddate.before(taskfilterenddate)) {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					} else if (reciveddate.equals(taskfilterstartdate) || reciveddate.equals(taskfilterenddate)) {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					}

				} else {

					clientforexchange_output.createTask(taskuri, task);

					count_destination++;

				}
			} else {

				if (task_box.isSelected()) {

					if (reciveddate.after(taskfilterstartdate) && reciveddate.before(taskfilterenddate)) {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					} else if (reciveddate.equals(taskfilterstartdate) || reciveddate.equals(taskfilterenddate)) {

						clientforexchange_output.createTask(taskuri, task);

						count_destination++;

					}

				} else {

					clientforexchange_output.createTask(taskuri, task);

					count_destination++;

				}
			}

		}

	}

	public void messageaddpst(MessageInfoCollection messageInfoCollection, String mailbox, String Folder) {
		String scm = "";
		int messagesize;
		if (demo) {
			if (messageInfoCollection.size() <= 50) {
				messagesize = messageInfoCollection.size();
			} else {
				messagesize = 50;
			}

		} else {
			messagesize = messageInfoCollection.size();
		}

		for (int i = 0; i < messagesize; i++) {
			try {

				if (stop) {
					break;
				}

				if ((i % 100) == 0) {
					System.gc();

				}

				MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MailConversionOptions de = new MailConversionOptions();
				de.setConvertAsTnef(true);
				MailMessage mess = message1.toMailMessage(de);
				MapiMessage message = MapiMessage.fromMailMessage(mess, d);

				Date reciveddate = message.getDeliveryTime();
				if (message.getMessageClass().equals("IPM.Contact")) {
					try {

						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };

						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getContactsUri(), calendertime + "/" + Folder,
								subfolderInfo)) {
							scm = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
											calendertime + "/" + Folder, null, "IPF.Contact")
									.getUri();
						}

						MapiContact con = (MapiContact) message.toMapiMessageItem();

						clientforexchange_output.createContact(scm, con);

						count_destination++;

					} catch (Error e) {
						mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
					} catch (Exception e) {
						mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

					try {

						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

						ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
						if (!clientforexchange_output.folderExists(
								clientforexchange_output.getMailboxInfo().getCalendarUri(), calendertime + "/" + Folder,
								subfolderInfo)) {
							scm = clientforexchange_output
									.createFolder(clientforexchange_output.getMailboxInfo().getCalendarUri(),
											calendertime + "/" + Folder, null, "IPF.Appointment")
									.getUri();
						}

						cal.save(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics",
								AppointmentSaveFormat.Ics);
						File file = new File(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics");
						AppointmentLoadOptions optiona = new AppointmentLoadOptions();
						optiona.getIgnoreSmtpAddressCheck();
						Appointment calDoc = Appointment
								.load(temppathm + File.separator + mf.namingconventionmapi(message) + ".ics", optiona);

						if (chckbx_calender_box.isSelected()) {
							if (reciveddate.after(Calenderfilterstartdate)
									&& reciveddate.before(Calenderfilterenddate)) {
								clientforexchange_output.createAppointment(calDoc, scm);
								count_destination++;
							}

						} else {
							clientforexchange_output.createAppointment(calDoc, scm);
							count_destination++;

						}
						file.delete();

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception :" + e.getMessage() + "  " + mf.namingconventionmapi(message)
								+ System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (messageInfo.getMessageClass().equals("IPM.StickyNote")
						|| messageInfo.getMessageClass().equals("IPM.Task")) {
					try {

						Boolean checktask = false;
						MapiTask task = null;
						if (messageInfo.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message.toMapiMessageItem();
						}
						String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();
						psttask(task, reciveddate, message, taskuri, checktask, messageInfo);

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception :" + e.getMessage() + "Task" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else {

					try {

						mailexchange(mess, reciveddate, clientforexchange_output, mailbox, "");

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception :" + e.getMessage() + "Message" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				}

				lbl_progressreport.setText("  Total Message Saved Count  " + count_destination + "  " + Folder
						+ "   Extarcting messsage " + message.getSubject());

			} catch (Exception e) {
				continue;
			}

		}
	}

	public void getsubfolderforPSTOST_defaultrestore(FolderInfo f, String path2, String path1, String mailfolder) {
		FolderInfoCollection subfolder = f.getSubFolders();
		String scm = "";
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (stop) {
					break;
				}

				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName().replaceAll("[\\[\\]]", "");

				path = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (stop) {
						break;
					}
					String path3 = pstfolderlist.get(l).replace(path2 + File.separator, "");
					if (path3.equalsIgnoreCase(path)) {

						String subfolder1 = clientforexchange_output.createFolder(mailfolder, Folder).getUri();
						// //int count = 1;
						lbl_progressreport.setText(" Getting Folder " + Folder);
						if (folderf.getContainerClass().contains("IPF.Appointment")) {

							if (folderf.getContentCount() > 0) {

								MessageInfoCollection messageInfoCollection = folderf.getContents();
								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {

										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										if (message.getMessageClass().equals("IPM.Appointment")
												|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {
												MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
												ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
												if (!clientforexchange_output.folderExists(
														clientforexchange_output.getMailboxInfo().getCalendarUri(),
														calendertime + "/" + Folder, subfolderInfo)) {
													scm = clientforexchange_output.createFolder(
															clientforexchange_output.getMailboxInfo().getCalendarUri(),
															calendertime + "/" + Folder, null, "IPF.Appointment")
															.getUri();
												}
												cal.save(temppathm + File.separator + mf.namingconventionmapi(message)
														+ ".ics", AppointmentSaveFormat.Ics);
												File file = new File(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics");
												AppointmentLoadOptions optiona = new AppointmentLoadOptions();
												optiona.getIgnoreSmtpAddressCheck();
												Appointment calDoc = Appointment.load(temppathm + File.separator
														+ mf.namingconventionmapi(message) + ".ics", optiona);
												clientforexchange_output.createAppointment(calDoc, scm);
												count_destination++;
												file.delete();
											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "  "
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}

							}
						} else if (folderf.getContainerClass().contains("IPF.Task")
								|| folderf.getContainerClass().contains("IPF.StickyNote")) {

							if (folderf.getContentCount() > 0) {

								MessageInfoCollection messageInfoCollection = folderf.getContents();
								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {

										if (stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										String taskuri = clientforexchange_output.getMailboxInfo().getTasksUri();

										Date reciveddate = message.getDeliveryTime();
										if (messageInfo.getMessageClass().equals("IPM.StickyNote")
												|| messageInfo.getMessageClass().equals("IPM.Task")) {
											try {

												MapiTask task = null;
												Boolean checktask = false;
												if (messageInfo.getMessageClass().equals("IPM.Task")) {
													task = (MapiTask) message.toMapiMessageItem();
												}

												if (chckbxRemoveDuplicacy.isSelected()) {
													String input = "";
													if (messageInfo.getMessageClass().equals("IPM.Task")) {
														input = mf.duplicacymapiTask(task);
													}

													if (!listduplictask.contains(input)) {
														System.out.println("Not a duplicate message");
														listduplictask.add(input);

														if (task_box.isSelected()) {
															if (reciveddate.after(taskfilterstartdate)
																	&& reciveddate.before(taskfilterenddate)) {
																psttask(task, reciveddate, message, taskuri, checktask,
																		messageInfo);
																count_destination++;
															} else if (reciveddate.equals(taskfilterstartdate)
																	|| reciveddate.equals(taskfilterenddate)) {
																psttask(task, reciveddate, message, taskuri, checktask,
																		messageInfo);
																count_destination++;

															}

														} else {
															psttask(task, reciveddate, message, taskuri, checktask,
																	messageInfo);
															count_destination++;

														}

													}
												} else {

													if (task_box.isSelected()) {
														if (reciveddate.after(taskfilterstartdate)
																&& reciveddate.before(taskfilterenddate)) {
															psttask(task, reciveddate, message, taskuri, checktask,
																	messageInfo);
															count_destination++;
														} else if (reciveddate.equals(taskfilterstartdate)
																|| reciveddate.equals(taskfilterenddate)) {
															psttask(task, reciveddate, message, taskuri, checktask,
																	messageInfo);
															count_destination++;

														}

													} else {
														psttask(task, reciveddate, message, taskuri, checktask,
																messageInfo);
														count_destination++;

													}

												}

											} catch (OutOfMemoryError ep) {
												mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
														+ mf.namingconventionmapi(message));
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Task" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else if (folderf.getContainerClass().contains("IPF.Contact")) {

							// System.out.println(Folder + "Folder Name");
							if (folderf.getContentCount() > 0) {

								ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
								if (!clientforexchange_output.folderExists(
										clientforexchange_output.getMailboxInfo().getContactsUri(),
										calendertime + "/" + Folder, subfolderInfo)) {
									scm = clientforexchange_output
											.createFolder(clientforexchange_output.getMailboxInfo().getContactsUri(),
													calendertime + "/" + Folder, null, "IPF.Contact")
											.getUri();
								}

								MessageInfoCollection messageInfoCollection = folderf.getContents();

								int messagesize;
								if (demo) {
									if (messageInfoCollection.size() <= 50) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = 50;
									}

								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {

										if (stop) {
											break;
										}

										if ((i % 100) == 0) {
											System.gc();

										}

										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
										MailConversionOptions de = new MailConversionOptions();
										de.setConvertAsTnef(true);
										MailMessage mess = message1.toMailMessage(de);
										MapiMessage message = MapiMessage.fromMailMessage(mess, d);

										if (message.getMessageClass().equals("IPM.Contact")) {
											try {

												MapiContact con = (MapiContact) message.toMapiMessageItem();
												if (chckbxRemoveDuplicacy.isSelected()) {

													String input = mf.duplicacymapiContact(con);

													if (!listdupliccontact.contains(input)) {

														listdupliccontact.add(input);

														clientforexchange_output.createContact(scm, con);
														count_destination++;

													} else {
														// System.out.println(" duplicate message");
														// System.out.println(input);

													}
												} else {

													clientforexchange_output.createContact(scm, con);
													count_destination++;

												}

												count_destination++;

											} catch (Error e) {
												mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
											} catch (Exception e) {
												mf.logger.warning("Exception :" + e.getMessage() + "Contact" + " " + i
														+ mf.namingconventionmapi(message) + System.lineSeparator());
												e.printStackTrace();
												continue;
											}

										}

										lbl_progressreport.setText("  Total Message Saved Count  " + count_destination
												+ "  " + Folder + "   Extarcting messsage " + message.getSubject());

									} catch (Exception e) {
										continue;
									}

								}
							}
						} else {

							if (folderf.getContentCount() > 0) {
								MessageInfoCollection messageInfoCollection = folderf.getContents();

								messageaddpst(messageInfoCollection, subfolder1, Folder);
							}
						}
						listdupliccal.clear();
						listduplicacy.clear();
						listdupliccontact.clear();
						listduplictask.clear();
						if (folderf.hasSubFolders()) {
							getsubfolderforPSTOST_defaultrestore(folderf, path2, path, subfolder1);
						}

					}
				}

				path = path.replace(File.separator + Folder, "");
			} catch (Exception e) {
				continue;
			}

		}
	}

	private void ConvertMbox_365(String Foldername) {
		FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);
		MboxrdStorageReader mbox = new MboxrdStorageReader(stream.toInputStream(), new MboxLoadOptions());
		MailMessage message = mbox.readNextMessage();



		listduplicacy.clear();
		while (!(message == null)) {
			try {
				if (stop) {
					break;
				}

				if (demo) {
					if (count_destination == 50)
						break;
				}

				String messagsub = message.getSubject();
				try {

					Date reciveddate = message.getDate();
					mailexchange(message, reciveddate, clientforexchange_output, Folderuri, Foldername);

				} catch (Exception e) {

					if (e.getMessage()
							.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
							|| e.getMessage().contains(
									"No connection could be made because the target machine actively refused it.")
							|| e.getMessage().contains("Rate limit hit")

							|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
									"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
						Progressbar.setVisible(false);
						connectionHandle(e.getMessage());

					} else {
						message = mbox.readNextMessage();
					}

					mf.logger.warning("Exception  " + e.getMessage() + "Message  Subject : " + messagsub
							+ System.lineSeparator());
					continue;
				}

				try {
					lbl_progressreport.setText("<html><b>Total Message Saved Count " + count_destination
							+ " Extarcting messsage " + message.getSubject());
					// messageprogress = (count * 100) / messagesize;
					// iprogressBar_message_p3.setValue(messageprogress);
					message.close();
					message = mbox.readNextMessage();
				} catch (Exception e) {

					mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());
					continue;
				}
				//// iprogressBar_message_p3.setValue(messageprogress);
				// count++;
				// count_destination++;
				if ((count_destination % 100) == 0) {
					System.gc();
				}
			} catch (Exception e) {
				continue;
			}

		}
		stream.close();
		mbox.dispose();

	}

	private void ConvertMbox_defaultrestore() {
		FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);
		MboxrdStorageReader mbox = new MboxrdStorageReader(stream.toInputStream(), new MboxLoadOptions());
		MailMessage message = mbox.readNextMessage();

		if (comboBox_1.getSelectedItem().toString().equalsIgnoreCase("Inbox")) {

			Folderuri = clientforexchange_output.getMailboxInfo().getInboxUri();

		} else if (comboBox_1.getSelectedItem().toString().equalsIgnoreCase("Sent Item")) {
			Folderuri = clientforexchange_output.getMailboxInfo().getDeletedItemsUri();

		} else if (comboBox_1.getSelectedItem().toString().equalsIgnoreCase("Draft")) {
			Folderuri = clientforexchange_output.getMailboxInfo().getDraftsUri();
		}

		listduplicacy.clear();
		while (!(message == null)) {
			try {
				if (stop) {
					break;
				}
				if (demo) {
					if (count_destination == 50)
						break;
				}

				String messagsub = "";
				try {
					messagsub = message.getSubject();
				} catch (Exception e1) {

				}
				try {

					Date reciveddate = message.getDate();
					mailexchange(message, reciveddate, clientforexchange_output, Folderuri, "");

				} catch (Exception e) {

					if (e.getMessage()
							.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
							|| e.getMessage().contains(
									"No connection could be made because the target machine actively refused it.")
							|| e.getMessage().contains("Rate limit hit")

							|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
									"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
						Progressbar.setVisible(false);
						connectionHandle(e.getMessage());

					} else {
						message = mbox.readNextMessage();
					}

					mf.logger.warning("Exception  " + e.getMessage() + "Message  Subject : " + messagsub
							+ System.lineSeparator());
					continue;
				}

				try {
					lbl_progressreport.setText("<html><b>Total Message Saved Count " + count_destination
							+ " Extarcting messsage " + message.getSubject());
					message.close();
					message = mbox.readNextMessage();
				} catch (Exception e) {

					mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());
					continue;
				}

				if ((count_destination % 100) == 0) {
					System.gc();
				}
			} catch (Exception e) {
				continue;
			}

		}
		stream.close();
		mbox.dispose();

	}

	public void connectionHandle(String gotMessage) {
		lbl_progressreport.setText("INTERNET Connection  LOST ");

		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			try {
				lbl_progressreport.setText("Connecting to Server Please Wait");
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				}

				else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				}

				else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				lbl_progressreport.setText("Connection extablished Retriving Messasge");
				break;
			} catch (Exception e) {
				lbl_progressreport.setText("INTERNET Connection  LOST ");

			}

		}

		Progressbar.setVisible(true);

	}

	public boolean isValid(String email) {
		String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\." + "[a-zA-Z0-9_+&*-]+)*@" + "(?:[a-zA-Z0-9-]+\\.)+[a-z"
				+ "A-Z]{2,7}$";

		Pattern pat = Pattern.compile(emailRegex);
		if (email == null)
			return false;
		return pat.matcher(email).matches();
	}

	public void connectionHandle1(String gotMessage) {
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));
		while (true) {
			try {

				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				}

				else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				}

				else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				break;

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

//	@SuppressWarnings("unchecked")
	public static void visitAllNodes(DefaultMutableTreeNode roe) {

		Enumeration<TreeNode> e = roe.depthFirstEnumeration();
		while (e.hasMoreElements()) {
			DefaultMutableTreeNode node = (DefaultMutableTreeNode) e.nextElement();

			lists.add(node);
			listst.add(node.toString().replace("<html><b>", ""));

		}

	}

	public IEWSClient connectionwithexchangeserver_output() throws Exception {

		clientforexchange_output = EWSClient.getEWSClient("https://" + domain_p3 + "/ews/Exchange.asmx", username_p3,
				password_p3);

		clientforexchange_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;

	}

	public IEWSClient conntiontohotmail_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p3,
				password_p3);

		clientforexchange_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	public ImapClient connectiontoimap_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();

		clientforimap_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_output;
	}

	public ImapClient connectiontogmail_output() throws Exception {
		clientforimap_output = new ImapClient("imap.gmail.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();
		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_output;
	}

	public ImapClient connectiontoaol_output() throws Exception {
		clientforimap_output = new ImapClient("imap.aol.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();
		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_output;
	}

	public ImapClient connectiontoyahoo_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.yahoo.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();
		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_output;
	}

	public IEWSClient conntiontooffice365_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);

		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5*60*1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	private static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName;
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "").replace("|", "")
				.replace("*", "").replace("<", "").replace(">", "").replace("\t", "").replace("\"", "")
				.replace(",", "");

		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}

	public static void search(TreeNode rootNode, DefaultMutableTreeNode searchNode) {

		for (int i = 0; i < rootNode.getChildCount(); i++) {

			if (rootNode.getChildAt(i).toString().equals(searchNode.toString())) {

				lastNode = (DefaultMutableTreeNode) rootNode.getChildAt(i);
			}

			else {

				search((DefaultMutableTreeNode) rootNode.getChildAt(i), searchNode);
			}

		}

	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {

			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {

			}
		}
	}

	private static String getFileExtension(File file) {
		String fileName = file.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	private StringBuilder Duration(long startTime) {
		long elapsedTime = System.currentTimeMillis() - startTime;
		long elapsedSeconds = elapsedTime / 1000;
		long secondsDisplay = elapsedSeconds % 60;
		long elapsedMinutes = elapsedSeconds / 60;
		StringBuilder br = new StringBuilder();

		br.append(elapsedMinutes);
		br.append(':');
		br.append(secondsDisplay);
		return br;
	}

	private String filepath(File file) {
		String fileName = file.getAbsolutePath();
		String filepath = fileName.replace(file.getName(), "");
		return filepath;
	}

	String mapiexchange(MapiMessage message, Date reciveddate, IEWSClient clientforexchange_output1, String Folderuri) {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output1.appendMessage(Folderuri, message, true);
				foldermessagecount++;
				count_destination++;

			}
		}
		return Messageid;
	}

	String mailexchange(MailMessage message, Date reciveddate, IEWSClient clientforexchange_output1, String Folderuri,
			String Mailboxname) throws Exception {
		String Messageid = "";

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

						if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
							clientforexchange_output1.archiveItem(Folderuri, Messageid);
						}
						count_destination++;

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

						if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
							clientforexchange_output1.archiveItem(Folderuri, Messageid);
						}
						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
					if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
						clientforexchange_output1.archiveItem(Folderuri, Messageid);
					}
					count_destination++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
					if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
						clientforexchange_output1.archiveItem(Folderuri, Messageid);
					}

					count_destination++;

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
					if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
						clientforexchange_output1.archiveItem(Folderuri, Messageid);
					}
					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output1.appendMessage(Folderuri, message);
				if (Mailboxname.equalsIgnoreCase("Archive Folder")) {
					clientforexchange_output1.archiveItem(Folderuri, Messageid);

				}
				count_destination++;

			}
		}
		return Messageid;
	}

	public void readDataLineByLine(String file) {
		FileReader filereader = null;
		CSVReader csvReader = null;
		try {
			DefaultTableModel model1 = (DefaultTableModel) table_userid.getModel();
			model1.setRowCount(0);
			HashSet<String> dupliset = new HashSet<String>();
			dupliset.clear();

			filereader = new FileReader(file);

			csvReader = new CSVReader(filereader);
			String[] nextRecord;

			while ((nextRecord = csvReader.readNext()) != null) {
				// System.out.println(nextRecord[0]);
				String email = nextRecord[0];

				String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\." + "[a-zA-Z0-9_+&*-]+)*@" + "(?:[a-zA-Z0-9-]+\\.)+[a-z"
						+ "A-Z]{2,7}$";

				Pattern pat = Pattern.compile(emailRegex);
				if (email != null) {
					if (pat.matcher(email).matches()) {
						if (dupliset.add(email)) {
							model1.addRow(new Object[] { true, "", email });
						}
					}
				}
				// System.out.println();
			}
		} catch (Exception e) {
			mf.logger.warning(e.getMessage() + System.lineSeparator());
		} finally {

			try {
				if (csvReader != null)
					csvReader.close();
				if (filereader != null)
					filereader.close();
			} catch (IOException e) {
				mf.logger.warning(e.getMessage() + System.lineSeparator());
			}
		}
	}

	private String bytes2String(long sizeInBytes) {
		NumberFormat nf = new DecimalFormat();
		nf.setMaximumFractionDigits(2);

		try {
			if (sizeInBytes < SPACE_KB) {
				return nf.format(sizeInBytes) + " Byte(s)";
			} else if (sizeInBytes < SPACE_MB) {
				return nf.format(sizeInBytes / SPACE_KB) + " KB";
			} else if (sizeInBytes < SPACE_GB) {
				return nf.format(sizeInBytes / SPACE_MB) + " MB";
			} else if (sizeInBytes < SPACE_TB) {
				return nf.format(sizeInBytes / SPACE_GB) + " GB";
			} else {
				return nf.format(sizeInBytes / SPACE_TB) + " TB";
			}
		} catch (Exception e) {
			return sizeInBytes + " Byte(s)";
		}
	}

}
