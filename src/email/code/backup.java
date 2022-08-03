package email.code;

import java.awt.AWTException;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.MenuItem;
import java.awt.PopupMenu;
import java.awt.SystemColor;
import java.awt.SystemTray;
import java.awt.Toolkit;
import java.awt.TrayIcon;
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
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Executors;
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
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;

import com.aspose.email.Appointment;
import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.Attachment;
import com.aspose.email.Contact;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.EWSClient;
import com.aspose.email.EmailAddress;
import com.aspose.email.EmailAddressList;
import com.aspose.email.EmailClient;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderInfoCollection;
import com.aspose.email.ExchangeMailboxInfo;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.ExchangeMessageInfoCollection;
import com.aspose.email.ExchangeMessagePageInfo;
import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapFolderInfo;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiMessageFlags;
import com.aspose.email.MapiTask;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MessageInfoType;
import com.aspose.email.OlmFolder;
import com.aspose.email.OlmStorage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SaveOptions;
import com.aspose.email.SecurityOptions;
import com.aspose.email.StandardIpmFolder;
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

@SuppressWarnings("deprecation")
public class backup extends JFrame {

	/**
	 * 
	 */
	static File licFileon;
	static File licFileonoff;
	TableColumnModel csvcolumnModel;
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	static JFileChooser jFileChooser;
	static JFileChooser jFileChooser_destination;
	static PersonalStorage pst;
	static String destination_path;
	String Folderuri;
	int kl = 0;
	public static String projectTitle;
	long countf = 0;
	int treefolder = 0;
	loadingThreadclassformailbox obThmailbox;
	String thunderbirdpath = "";
	ArrayList<String> pstfolderlist2;
	JPanel panel_progress;
	JCheckBox chckbxIMAdmin;
	TableColumnModel csvcolumnModel2;
	JLabel lblEnableImap_p3;
	HashMap<String, List<String>> hm;
	HashMap<String, List<String>> hmstring;
	HashMap<String, List<ExchangeFolderInfo>> hmexpo;
	HashMap<String, List<ExchangeFolderInfo>> hmexpofinal;
	Boolean contatcheck;
	String userid = "";
	JTextField textField_customfolder;
	JCheckBox chckbxCustomFolderName;
	Boolean calendarcheck;
	LoadingThreadclass obTh;
	JLabel lblTurnOffTwo_p3;
	JPanel panel_taskfilter;
	JCheckBox chckbxShowPassword_p2;
	List<String> pstfolderlist;
	static File file;
	int portmo;
	JLabel lblMakeSureYou;
	JLabel lblLive_Chat_p3;
	JButton button_stop;
	JCheckBox chckbxRemoveDuplicacy;
	JLabel lblPortNo;
	int portnofiletype;
	int ret;
	int doubleclickcount = 0;
	Date mailfilterstartdate;
	Date mailfilterenddate;
	Date Calenderfilterstartdate;
	Date Calenderfilterenddate;
	Date taskfilterstartdate;
	Date taskfilterenddate;
	String date;
	String demoversion = "Demo";
	String Fullversion = "Full";
	static String filefoldername;
	String fileaddress;
	private JButton btn_next_pane2;
	String filepath;
	static FolderInfo folderInfo;
	static String[] si;
	static String foldername = "";
	static String reportpath;
	static String fname;
	JPanel panel_1;
	JLabel Progressbar;
	JPanel panel_2;
	JPanel panel_2_2_1;
	JPanel panel_3;
	JPanel panel_3_;
	JPanel panel_3_1_2;
	JPanel panel_3_1_1;
	JPanel panel_3_1_1_1;
	JPanel panel_3_2;
	JPanel panel_Loading;
	JLabel label_8;

	int countcheck = 0;
	JPanel panel_computername;
	String Status = "Completed";
	JPanel panel_4;
	JPanel panel_mailfilter;
	Boolean checkdestination = true;
	Boolean checkconvertagain = false;
	String fileoption;
	String filetype = "Not selected";
	static CheckboxTree tree;
//	static JTree tree;
	static DefaultTableModel modeli;
	static DefaultTreeModel model;
	static DefaultMutableTreeNode root;
	static DefaultTableModel mode;
	static DefaultMutableTreeNode lastNode;
	JDateChooser dateChooser_mail_fromdate;
	JDateChooser dateChooser_mail_tilldate;

	static JComboBox<String> comboBox_fileDestination_type;
	JComboBox<String> comboBox;

	static JTextField tf_Destination_Location;
	private JPanel Cardlayout;
	private static JTable table_fileConvertionreport_panel4;
	static IEWSClient clientforexchange_output;
	static IEWSClient clientforexchange_input;
	static IConnection iconnforimap_input;
	static ImapClient clientforimap_input;
	static IConnection iconnforimap_output;
	static ImapClient clientforimap_output;
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	static String username_p2 = "";
	static String domain_p2 = "";
	static String password_p2 = "";
	static String username_p3 = "";
	static String domain_p3 = "";
	static String password_p3 = "";
	static long count_destination;

	String filetemp = "";
	ImapFolderInfo imapFolderInfo;

	static String messageboxtitle;
	String buyurl;
	String infourl;
	String helpurl;
	boolean stop = false;
	boolean stop_tree = false;
	static backup frame;
	static String calendertime;
	Boolean checkBackupRest = false;
	JLabel lbl_header;
	JLabel lbl_connecting_p2;
	JLabel lbl_DomainName_computername;
	static JLabel lbl_progressreport;
	private JTextField textField_username_p2;
	private JPasswordField passwordField_p2;
	private JTextField textField_Domainname_p2;
	static JCheckBox chckbx_Mail_Filter;
	static JCheckBox chckbx_calender_box;
	static Calendar cal;
	static String parentname = "";
	static String path = "";
	String path1 = "";
	static long foldermessagecount = 0;

	static OlmStorage storage;
	private JButton btnDowloadReport;
	private JTextField textField_username_p3;
	private JPasswordField passwordField_p3;
	static Iterator<MapiMessage> it;
	private JTextField textField_domain_name_p3;
	static String Folder;
	static String hostName = "";
	JButton btn_Destination;
	JButton btn_Sign_p3;
	JButton btn_Previous_pane2;
	JButton btn_converter;
	JButton btn_previous_p3;
	JButton btn_signout_p3;

	JButton btnStop;
	File f;
	Thread th;

	JPanel panel_Calender;
	static File index;
	private JCheckBox chckbxShowPassword_p3;
	private JLabel lbl_connecting_p3;
	static String mess;
	FolderInfo info1;
	static JProgressBar progressBar_message_p3;
	private JPanel panel_loginpanel;
	JPanel panel_mailfilt7;
	JLabel lbl_Domain;
	JDateChooser dateChooser_task_end_date;
	private JPanel inner_cardlayout;
	FolderInfo info = new FolderInfo();
	PersonalStorage ost;
	ExchangeFolderInfo exchangeFolderInfo;
	OlmFolder folderi;
	JCheckBox task_box;
	JDateChooser dateChooser_task_start_date;
	private final ButtonGroup buttonGroup = new ButtonGroup();
	private JLabel lblNewLabel_2;
	private JLabel lblReadingFoldersPlease;
	private JLabel lblNewLabel_3;
	private JLabel lblNewLabel_4;
	Boolean input = false;
	Boolean output = false;
	static Boolean demo = true;
	JButton btnSavingLog;
	String[] usernameverification;
	private JButton btnConvertAgain;
	private JLabel label_6;
	private JLabel label_9;
	private JLabel label_10;
	private List<MailMessage> listmail = new ArrayList<MailMessage>();
	private List<MapiMessage> listmapi = new ArrayList<MapiMessage>();
	String multifatcorauthicationfor365 = "";
	private List<String> userList = new ArrayList<String>();

	private List<String> listFolderinfostring = new ArrayList<String>();
	private List<ExchangeFolderInfo> listExchangemesingos = new ArrayList<ExchangeFolderInfo>();
	private List<ExchangeFolderInfo> listExchangdinal = new ArrayList<ExchangeFolderInfo>();
	private List<String> listduplicacy = new ArrayList<String>();
	private JLabel label_13;
	Boolean Stoppreview = false;
	public String logpath = "";
	String temppath = "";
	private JTextField tf_portNo_p3;
	private JLabel lblLiveChat;

	JLabel lblTurnOffTwo;
	private JPanel panel;
	private JLabel label_12;
	private JTable table_userid;
	private JTextField textField_3;
	String version;
	Main_Frame mf;
	private JButton btn_previous;

	private JPanel panel_9;
	private JRadioButton rdbtnMailbox;
	private JRadioButton rdbtnArchieve;
	private JRadioButton rdbtnPublicFolder;
	private JLabel label_15;
	private JLabel lblNewLabel_7;
	private JLabel lblNewLabel_8;
	private JLabel lblNewLabel_9;

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public backup(JFrame parent1, Boolean demo) {
		super();
		mf = (Main_Frame) parent1;
		backup.demo = demo;
		fileoption = Main_Frame.fileoption;
		messageboxtitle = All_Data.messageboxtitle;
		buyurl = mf.buyurl;
		infourl = mf.infourl;
		helpurl = mf.helpurl;
		logpath = mf.logpath;
		temppath = mf.temppath;
		version = mf.version;
		projectTitle = All_Data.messageboxtitle;
		multifatcorauthicationfor365 = mf.multifatcorauthicationfor365;

		addWindowListener(new WindowAdapter() {

			public void windowClosing(WindowEvent arg0) {

				if (!SystemTray.isSupported()) {
					String warn = "Do you want to close the application?";
					int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
					if (ans == JOptionPane.YES_OPTION) {

						openBrowser(infourl);
						System.exit(0);
					}

				}

				SystemTray systemTray = SystemTray.getSystemTray();

				PopupMenu trayPopupMenu = new PopupMenu();
				MenuItem action = new MenuItem("Show");
				action.addActionListener(new ActionListener() {
					@Override
					public void actionPerformed(ActionEvent e) {

						setVisible(true);
					}
				});
				trayPopupMenu.add(action);

				MenuItem close = new MenuItem("Exit");
				close.addActionListener(new ActionListener() {
					@Override
					public void actionPerformed(ActionEvent e) {

						String warn = "Do you want to close the application?";
						int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
								JOptionPane.INFORMATION_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
						if (ans == JOptionPane.YES_OPTION) {

							if (demo) {

								openBrowser(infourl);

							}

							System.exit(0);
						}

					}
				});
				trayPopupMenu.add(close);

				TrayIcon trayIcon = new TrayIcon(
						Toolkit.getDefaultToolkit().getImage(Main_Frame.class.getResource("/128x128.png")),
						messageboxtitle, trayPopupMenu);
				trayIcon.setImageAutoSize(true);

				try {

					TrayIcon[] icons = (TrayIcon[]) SystemTray.getSystemTray().getTrayIcons();

					boolean check = false;
					for (int i = 0; i < icons.length; i++) {

						if (icons[i].getImage().equals(trayIcon.getImage())) {
							check = true;
							break;
						}
					}

					if (!check) {
						systemTray.add(trayIcon);
						trayIcon.displayMessage("Tool Added in Tray ", " ", TrayIcon.MessageType.NONE);
					} else {
						System.out.println("tool already in tray");
					}
				} catch (AWTException awtException) {
					awtException.printStackTrace();
				}
				System.out.println("end of main");
				setVisible(false);

			}
		});

		setIconImage(Toolkit.getDefaultToolkit().getImage(backup.class.getResource("/128x128.png")));

//		if (demo) {
//			String center = messageboxtitle + " v" + version + "(" + mf.demoversion + ")";
//			
//		} else {
//			String center = messageboxtitle + " v" + version + "(" + mf.Fullversion + ")";
//			setTitle(center);
//		}
		if (demo) {

			String title = messageboxtitle.replace("(Enterprise License)", " ");
			String center = title + "(" + demoversion + ")";
			messageboxtitle = center;
			setTitle(center);
		} else {
			String center = messageboxtitle + "(" + Fullversion + ")";
			setTitle(center);
		}

		setLocationRelativeTo(null);
		setResizable(false);

		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		setBounds(100, 100, 1081, 718);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

//		JMenuBar menuBar = mf.menuBar;
//		
//		setJMenuBar(menuBar);

		Cardlayout = new JPanel();
		Cardlayout.setBackground(Color.LIGHT_GRAY);
		Cardlayout.setBounds(0, 69, 1075, 620);
		contentPane.add(Cardlayout);
		Cardlayout.setLayout(new CardLayout(0, 0));

		JPopupMenu menu = mf.menu;

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
		btnActivate.setToolTipText("Click here to activate the software.");
		btnActivate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				if (System.getProperty("os.name").toLowerCase().contains("windows")) {
					licFileon = new File(System.getenv("APPDATA") + File.separator + projectTitle + File.separator
							+ "licenseOnline.lic");

					licFileonoff = new File(
							System.getenv("APPDATA") + File.separator + projectTitle + File.separator + "license.lic");

				} else {
					licFileon = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
							+ "Application Support" + File.separator + projectTitle + File.separator
							+ "licenseOnline.lic");
					licFileonoff = new File(
							System.getenv("APPDATA") + File.separator + projectTitle + File.separator + "license.lic");
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

				messageboxtitle = projectTitle;

			}
		});

		btnActivate.setBounds(949, 18, 32, 40);
		contentPane.add(btnActivate);
		panel_1 = new JPanel();
		panel_1.setBackground(Color.WHITE);
		Cardlayout.add(panel_1, "panel_1");

		JButton btn_info = new JButton("");
		btn_info.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				AboutDialog ab;
				if (demo) {
					ab = new AboutDialog(frame, true, "demo");

				} else {
					ab = new AboutDialog(frame, true, "full");
				}
				ab.setLocationRelativeTo(frame);
				ab.setVisible(true);

			}
		});
		btn_info.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_info.setIcon(new ImageIcon(backup.class.getResource("/info-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_info.setIcon(new ImageIcon(backup.class.getResource("/info-btn.png")));
			}
		});

		JButton btn_buy = new JButton("");
		btn_buy.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openBrowser(buyurl);

			}
		});

		JButton btn_help = new JButton("");
		btn_help.setToolTipText("Click here for software guide.");
		btn_help.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(helpurl);
			}
		});
		btn_help.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_help.setIcon(new ImageIcon(backup.class.getResource("/about-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_help.setIcon(new ImageIcon(backup.class.getResource("/about-btn.png")));
			}
		});

		btn_help.setIcon(new ImageIcon(backup.class.getResource("/about-btn.png")));

		btn_help.setRolloverEnabled(false);
		btn_help.setRequestFocusEnabled(false);
		btn_help.setOpaque(false);
		btn_help.setFocusable(false);
		btn_help.setFocusTraversalKeysEnabled(false);
		btn_help.setFocusPainted(false);
		btn_help.setDefaultCapable(false);
		btn_help.setContentAreaFilled(false);
		btn_help.setBorderPainted(false);
		btn_help.setBounds(983, 17, 41, 41);
		contentPane.add(btn_help);
		btn_buy.setBounds(909, 18, 41, 41);
		btn_buy.setToolTipText("Click here to purchase the software.");
		contentPane.add(btn_buy);

		btn_buy.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_buy.setIcon(new ImageIcon(backup.class.getResource("/buy-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_buy.setIcon(new ImageIcon(backup.class.getResource("/buy-btn.png")));
			}
		});

		btn_buy.setIcon(new ImageIcon(backup.class.getResource("/buy-btn.png")));

		btn_buy.setOpaque(false);
		btn_buy.setRolloverEnabled(false);
		btn_buy.setRequestFocusEnabled(false);
		btn_buy.setFocusTraversalKeysEnabled(false);
		btn_buy.setFocusable(false);
		btn_buy.setFocusPainted(false);
		btn_buy.setDefaultCapable(false);
		btn_buy.setContentAreaFilled(false);
		btn_buy.setBorderPainted(false);
		if (demo) {
			btn_buy.setVisible(true);
		} else {
			btn_buy.setVisible(false);
		}

		btn_info.setIcon(new ImageIcon(backup.class.getResource("/info-btn.png")));
		btn_info.setRolloverEnabled(false);
		btn_info.setToolTipText("Click here for software information.");
		btn_info.setRequestFocusEnabled(false);
		btn_info.setOpaque(false);
		btn_info.setFocusable(false);
		btn_info.setFocusTraversalKeysEnabled(false);
		btn_info.setFocusPainted(false);
		btn_info.setDefaultCapable(false);
		btn_info.setContentAreaFilled(false);
		btn_info.setBorderPainted(false);
		btn_info.setBounds(1018, 17, 47, 41);
		contentPane.add(btn_info);
		panel_1.setLayout(null);

		JButton btn_previous_p1 = new JButton("");
		btn_previous_p1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous_p1.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous_p1.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});
		btn_previous_p1.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous_p1.setRolloverEnabled(false);
		btn_previous_p1.setRequestFocusEnabled(false);
		btn_previous_p1.setOpaque(false);
		btn_previous_p1.setFocusable(false);

		btn_previous_p1.setFocusTraversalKeysEnabled(false);
		btn_previous_p1.setFocusPainted(false);
		btn_previous_p1.setDefaultCapable(false);
		btn_previous_p1.setContentAreaFilled(false);
		btn_previous_p1.setBorderPainted(false);
		btn_previous_p1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Main_Frame main_m = new Main_Frame(demo, Main_Frame.versiontype);
				main_m.setVisible(true);
				main_m.setLocationRelativeTo(null);
				main_m.setResizable(false);
				main_m.rdbtnSingleFile.setSelected(true);
				dispose();
			}
		});
		btn_previous_p1.setFont(new Font("Tahoma", Font.BOLD, 12));
		btn_previous_p1.setBounds(942, 571, 123, 38);
		panel_1.add(btn_previous_p1);

		inner_cardlayout = new JPanel();
		inner_cardlayout.setBounds(408, 83, 657, 383);
		panel_1.add(inner_cardlayout);
		inner_cardlayout.setLayout(new CardLayout(0, 0));

		panel_loginpanel = new JPanel();
		panel_loginpanel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_loginpanel.setBackground(Color.WHITE);

		JLabel lblMakeSureYou_1 = new JLabel("Please  Click on The Link");
		lblMakeSureYou_1.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblMakeSureYou_1.setForeground(Color.BLACK);
		lblMakeSureYou_1.setBounds(10, 326, 174, 32);
		panel_loginpanel.add(lblMakeSureYou_1);
		// lblMakeSureYou.setFont(new Font("Tahoma", Font.BOLD, 13));

		JLabel lblTurnOffTwo_p1 = new JLabel(
				"If your multi-factor authentication is enabled, then you need to create third-party app password.");
		lblTurnOffTwo_p1.setBounds(161, 326, 496, 32);
		panel_loginpanel.add(lblTurnOffTwo_p1);
		lblTurnOffTwo_p1.setForeground(Color.BLUE);
		lblTurnOffTwo_p1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openBrowser(All_Data.multifatcorauthicationfor365);
			}
		});
		lblTurnOffTwo_p1.setCursor(mf.cursor);
		lblTurnOffTwo_p1.setFont(new Font("Tahoma", Font.PLAIN, 11));

		inner_cardlayout.add(panel_loginpanel, "panel_loginpanel");
		panel_loginpanel.setLayout(null);

		JLabel lbl_username = new JLabel("USER NAME");
		lbl_username.setBounds(10, 110, 146, 27);
		panel_loginpanel.add(lbl_username);
		lbl_username.setFont(new Font("Tahoma", Font.BOLD, 14));

		JLabel lbl_password = new JLabel("PASSWORD");
		lbl_password.setBounds(10, 148, 181, 27);
		panel_loginpanel.add(lbl_password);
		lbl_password.setFont(new Font("Tahoma", Font.BOLD, 14));

		textField_username_p2 = new JTextField();
		textField_username_p2.setHorizontalAlignment(JTextField.CENTER);
		textField_username_p2.setComponentPopupMenu(menu);
		textField_username_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		textField_username_p2.setBounds(201, 109, 298, 30);
		panel_loginpanel.add(textField_username_p2);
		textField_username_p2.setColumns(10);

		passwordField_p2 = new JPasswordField();
		passwordField_p2.setHorizontalAlignment(JTextField.CENTER);
		passwordField_p2.setComponentPopupMenu(menu);
		passwordField_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		passwordField_p2.setBounds(201, 150, 298, 30);
		panel_loginpanel.add(passwordField_p2);

		lbl_connecting_p2 = new JLabel("");
		lbl_connecting_p2.setBounds(336, 302, 83, 39);
		panel_loginpanel.add(lbl_connecting_p2);
		lbl_connecting_p2.setIcon(new ImageIcon(backup.class.getResource("/loading.gif")));

		lbl_connecting_p2.setVisible(false);

		chckbxShowPassword_p2 = new JCheckBox("Show Password");
		chckbxShowPassword_p2.setBackground(Color.WHITE);
		chckbxShowPassword_p2.setBounds(505, 150, 142, 25);
		panel_loginpanel.add(chckbxShowPassword_p2);
		chckbxShowPassword_p2.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField_p2.setEchoChar((char) 0);
				}

				else {
					passwordField_p2.setEchoChar('*');
				}
			}
		});
		chckbxShowPassword_p2.setFont(new Font("Tahoma", Font.BOLD, 13));

		JButton btn_SignIn = new JButton("");
		btn_SignIn.setRolloverEnabled(false);
		btn_SignIn.setRequestFocusEnabled(false);
		btn_SignIn.setOpaque(false);
		btn_SignIn.setFocusable(false);
		btn_SignIn.setFocusTraversalKeysEnabled(false);
		btn_SignIn.setFocusPainted(false);
		btn_SignIn.setDefaultCapable(false);
		btn_SignIn.setContentAreaFilled(false);
		btn_SignIn.setBorderPainted(false);
		btn_SignIn.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_SignIn.setIcon(new ImageIcon(backup.class.getResource("/sign-in-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_SignIn.setIcon(new ImageIcon(backup.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_SignIn.setIcon(new ImageIcon(backup.class.getResource("/sign-in-btn.png")));
		btn_SignIn.setBounds(513, 256, 118, 39);
		btn_SignIn.setToolTipText("Click here to Sign in to " + fileoption);
		panel_loginpanel.add(btn_SignIn);
		btn_SignIn.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				try {
					domain_p2 = textField_Domainname_p2.getText();

				} catch (Exception a) {

				}
				try {
					username_p2 = textField_username_p2.getText();
					username_p2 = username_p2.trim();

				} catch (Exception a) {

				}
				try {
					password_p2 = new String(passwordField_p2.getPassword());
					password_p2 = password_p2.trim();

				} catch (Exception a) {

				}
				try {

				} catch (Exception a) {

				}

				if (username_p2.equalsIgnoreCase("") || password_p2.equalsIgnoreCase("")) {

					if (username_p2.equalsIgnoreCase("") && password_p2.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(frame, "User name and Password fields cannot be empty",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(backup.class.getResource("/information.png")));

					} else if (username_p2.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "User name field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));

					} else if (password_p2.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "Password field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));

					}

					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					btn_signout_p3.setVisible(true);
					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);
				} else if (fileoption.equalsIgnoreCase("Exchange Backup & Restore") && domain_p2.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "Computer Name or IP Address field can not be empty",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (!isValid(username_p2)) {

					JOptionPane.showMessageDialog(frame, "Please enter a valid username", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));

					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else {
					th = new Thread(new Runnable() {

						public void run() {

							lbl_connecting_p2.setVisible(true);
							textField_Domainname_p2.setEnabled(false);
							passwordField_p2.setEnabled(false);
							chckbxIMAdmin.setEnabled(false);
							textField_username_p2.setEnabled(false);
							btnSavingLog.setEnabled(false);

							chckbxShowPassword_p2.setEnabled(false);
							btn_SignIn.setEnabled(false);

							if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {

								mailboxUri = "https://" + domain_p2 + "/ews/Exchange.asmx";
							}

							try {
								button_stop.setVisible(true);
								btn_Previous_pane2.setEnabled(false);
								btn_next_pane2.setEnabled(false);

								mf.logger = mf.logFile();
								cal = Calendar.getInstance();
								calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
								mf.logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : "
										+ fileoption + "                         " + "Mailbox" + "    "
										+ textField_username_p2.getText() + System.lineSeparator()
										+ "======================================================================");

								conntiontooffice365_input();
								usernameverification = username_p2.split("@");
								button_stop.setVisible(false);
								rdbtnMailbox.setSelected(true);
								cal = Calendar.getInstance();
								lbl_header.setText(fileoption);

								calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

								try {
									conntiontooffice365_input();
								} catch (Exception e) {

									e.printStackTrace();
								}

								// obThmailbox.ob.MessageLabel.setText("Getting Mailbox User Id Please
								// wait...");

								if (chckbxIMAdmin.isSelected()) {
									impersonationUser();
								} else {

									DefaultTableModel model = (DefaultTableModel) table_userid.getModel();

									model.addRow(new Object[] { true, "", username_p2, "Mailbox" });
								}

								CardLayout card = (CardLayout) Cardlayout.getLayout();
								card.show(Cardlayout, "panel_tablecontentBR");

								chckbxShowPassword_p2.setEnabled(true);
								chckbxIMAdmin.setEnabled(true);
							} catch (Error e) {
								mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								textField_Domainname_p2.setEnabled(true);
								passwordField_p2.setEnabled(true);
								// textField.setEnabled(true);
								btnSavingLog.setEnabled(true);
								// btnTempPath.setEnabled(true);
								lblLiveChat.setVisible(true);
								textField_username_p2.setEnabled(true);
								mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
								if (fileoption.equalsIgnoreCase("Gmail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Gmail please check your Credantial Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(backup.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(backup.class.getResource("/information.png")));
									}
								} else if (fileoption.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(backup.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {

									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(backup.class.getResource("/information.png")));
									}
								} else if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(frame, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(backup.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(frame, "Connection not established", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(backup.class.getResource("/information.png")));
								}

								return;

							} finally {
								lbl_connecting_p2.setVisible(false);
								btn_Previous_pane2.setEnabled(true);
								btn_next_pane2.setEnabled(true);
								passwordField_p2.setEditable(true);
								chckbxShowPassword_p2.setEnabled(true);
								passwordField_p2.setEnabled(true);
								btn_SignIn.setEnabled(true);
								stop_tree = false;

							}

						}
					});
					th.start();

				}
			}
		});
		btn_SignIn.setFont(new Font("Tahoma", Font.BOLD, 14));

		panel_computername = new JPanel();
		panel_computername.setBackground(Color.WHITE);
		panel_computername.setBounds(10, 201, 506, 39);
		panel_loginpanel.add(panel_computername);
		panel_computername.setLayout(null);

		lbl_DomainName_computername = new JLabel("");
		lbl_DomainName_computername.setBounds(0, 0, 184, 27);
		panel_computername.add(lbl_DomainName_computername);
		if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			lbl_DomainName_computername.setText("Domain Name");
			;

		}
		lbl_DomainName_computername.setFont(new Font("Tahoma", Font.BOLD, 10));

		textField_Domainname_p2 = new JTextField();
		textField_Domainname_p2.setHorizontalAlignment(JTextField.CENTER);
		textField_Domainname_p2.setComponentPopupMenu(menu);
		textField_Domainname_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		textField_Domainname_p2.setBounds(194, 0, 297, 32);
		panel_computername.add(textField_Domainname_p2);
		if (!fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			textField_Domainname_p2.setVisible(false);

		}

		textField_Domainname_p2.setColumns(10);

		label_8 = new JLabel("");
		label_8.setBounds(0, 0, 661, 68);
		panel_loginpanel.add(label_8);

		if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			panel_computername.setVisible(true);
		}

		lblLiveChat = new JLabel("More Help");
		lblLiveChat.setForeground(Color.RED);
		lblLiveChat.setFont(new Font("Tahoma", Font.PLAIN, 14));

		lblLiveChat.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");

			}
		});

		lblLiveChat.setBounds(509, 92, 76, 27);
		panel_loginpanel.add(lblLiveChat);

		lblTurnOffTwo = new JLabel("");
		lblTurnOffTwo.setForeground(Color.RED);
		lblTurnOffTwo.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openBrowser(multifatcorauthicationfor365);
			}
		});
		if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			lblTurnOffTwo.setVisible(false);
		}
		lblTurnOffTwo.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo.setBounds(175, 245, 472, 14);
		panel_loginpanel.add(lblTurnOffTwo);

		chckbxIMAdmin = new JCheckBox("Assign Impersonation right");
		chckbxIMAdmin.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxIMAdmin.setBackground(Color.WHITE);
		chckbxIMAdmin.setBounds(278, 272, 186, 23);
		panel_loginpanel.add(chckbxIMAdmin);

		lblNewLabel_8 = new JLabel("Enter Email Address");
		lblNewLabel_8.setForeground(Color.RED);
		lblNewLabel_8.setBounds(10, 135, 146, 14);
		panel_loginpanel.add(lblNewLabel_8);

		lblNewLabel_9 = new JLabel("Password/Third-Party App Password ");
		lblNewLabel_9.setForeground(Color.RED);
		lblNewLabel_9.setBounds(10, 170, 193, 27);
		panel_loginpanel.add(lblNewLabel_9);

		panel_Loading = new JPanel();
		panel_Loading.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_Loading.setBackground(Color.WHITE);
		inner_cardlayout.add(panel_Loading, "panel_Loading");
		panel_Loading.setLayout(null);

		lblNewLabel_2 = new JLabel("Login Successful");
		lblNewLabel_2.setForeground(SystemColor.textHighlight);
		lblNewLabel_2.setFont(new Font("Tahoma", Font.BOLD, 18));
		lblNewLabel_2.setBounds(218, 13, 318, 44);
		panel_Loading.add(lblNewLabel_2);

		lblReadingFoldersPlease = new JLabel("");
		lblReadingFoldersPlease.setText("Reading Folders Please wait......");
		lblReadingFoldersPlease.setFont(new Font("Tahoma", Font.BOLD, 17));
		lblReadingFoldersPlease.setBounds(25, 83, 279, 44);
		panel_Loading.add(lblReadingFoldersPlease);

		lblNewLabel_3 = new JLabel("");
		lblNewLabel_3.setFont(new Font("Tahoma", Font.BOLD, 10));
		lblNewLabel_3.setBounds(35, 124, 604, 41);
		panel_Loading.add(lblNewLabel_3);

		lblNewLabel_4 = new JLabel("");
		lblNewLabel_4.setIcon(new ImageIcon(backup.class.getResource("/progress-bar.gif")));
		lblNewLabel_4.setBounds(10, 240, 637, 30);
		panel_Loading.add(lblNewLabel_4);

		button_stop = new JButton("");

		button_stop.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button_stop.setIcon(new ImageIcon(backup.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				button_stop.setIcon(new ImageIcon(backup.class.getResource("/stop-btn.png")));
			}
		});

		button_stop.setIcon(new ImageIcon(backup.class.getResource("/stop-btn.png")));
		button_stop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(backup.this, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE, new ImageIcon(backup.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					stop_tree = true;
				}

			}
		});
		button_stop.setRolloverEnabled(false);
		button_stop.setRequestFocusEnabled(false);
		button_stop.setOpaque(false);

		button_stop.setFocusable(false);
		button_stop.setFocusTraversalKeysEnabled(false);
		button_stop.setFocusPainted(false);
		button_stop.setDefaultCapable(false);
		button_stop.setContentAreaFilled(false);
		button_stop.setBorderPainted(false);
		button_stop.setBounds(246, 281, 142, 38);
		panel_Loading.add(button_stop);

		label_6 = new JLabel("");
		label_6.setIcon(new ImageIcon(backup.class.getResource("/sidebar.png")));
		label_6.setBounds(0, 0, 397, 620);
		panel_1.add(label_6);

		btnSavingLog = new JButton("");

		JPanel panel_tablecontentBR = new JPanel();
		panel_tablecontentBR.setBackground(Color.WHITE);
		Cardlayout.add(panel_tablecontentBR, "panel_tablecontentBR");
		panel_tablecontentBR.setLayout(null);

		JScrollPane scrollPane_2 = new JScrollPane();
		scrollPane_2.setBounds(0, 77, 883, 472);
		panel_tablecontentBR.add(scrollPane_2);

		Object[][] data1 = {};

		String[] cols1 = { "", "", "Mailbox User Id", "" };

		DefaultTableModel tablemodel = new DefaultTableModel(data1, cols1) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {
				// all cells false
				if (column == 0) {
					return true;
				} else if (column == 3) {
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
		table_userid.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		DefaultTableCellRenderer tablerenderer = (DefaultTableCellRenderer) table_userid.getTableHeader()
				.getDefaultRenderer();
		tablerenderer.setHorizontalAlignment(0);
		csvcolumnModel = table_userid.getColumnModel();

		csvcolumnModel.getColumn(1).setCellRenderer(new ImageRenderer());
		csvcolumnModel.getColumn(1).setPreferredWidth(20);
		csvcolumnModel.getColumn(1).setMaxWidth(20);
		csvcolumnModel.getColumn(1).setHeaderRenderer(new ImageRenderer());
		csvcolumnModel.getColumn(0).setHeaderRenderer(new HeaderRenderer2(table_userid.getTableHeader(), 0));
		csvcolumnModel.getColumn(0).setPreferredWidth(30);
		csvcolumnModel.getColumn(0).setMaxWidth(30);
		csvcolumnModel.getColumn(2).setPreferredWidth(620);

		csvcolumnModel2 = table_userid.getColumnModel();
		TableColumn testColumn = csvcolumnModel2.getColumn(3);
		String[] destinationMailboxes = { "Mailbox", "Public Folder", "Archive Folder" };
		JComboBox jcomboxmail = new JComboBox(destinationMailboxes);
		testColumn.setCellEditor(new DefaultCellEditor(jcomboxmail));
		csvcolumnModel.getColumn(3).setHeaderRenderer(new ImageRenderer1());

		tablerenderer.setHorizontalAlignment(0);

		scrollPane_2.setViewportView(table_userid);

		JButton btnNext = new JButton("");
		btnNext.setBounds(923, 571, 142, 38);
		btnNext.setRolloverEnabled(false);
		btnNext.setRequestFocusEnabled(false);
		btnNext.setOpaque(false);

		JLabel label_14 = new JLabel("");
		label_14.setFont(new Font("Tahoma", Font.BOLD, 16));
		label_14.setText(fileoption);
		label_14.setBounds(434, 10, 597, 49);
		panel_1.add(label_14);

		JLabel label_11 = new JLabel("");
		label_11.setBounds(396, 563, 679, 57);
		label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		panel_1.add(label_11);
		btnNext.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Boolean tableuseridcheck = false;
				for (int i = 0; i < table_userid.getRowCount(); i++) {
					if ((boolean) table_userid.getValueAt(i, 0)) {
						tableuseridcheck = true;
						break;
					}
				}
				if (tableuseridcheck) {
					hmstring = new HashMap<String, List<String>>();
					hmexpo = new HashMap<String, List<ExchangeFolderInfo>>();

					SwingWorker sw1 = new SwingWorker() {
						@Override
						protected Object doInBackground() {

							obThmailbox = new loadingThreadclassformailbox(mf);
							obThmailbox.start();

							model = (DefaultTreeModel) tree.getModel();

							root = new DefaultMutableTreeNode("<html><b>" + fileoption);

							model.setRoot(root);

							for (int i = 0; i < table_userid.getRowCount(); i++) {
								String useridchecked;
								try {

									useridchecked = table_userid.getValueAt(i, 0).toString();

								} catch (Exception e1) {
									continue;
								}
								if (useridchecked.equalsIgnoreCase("True")) {
									CardLayout card = (CardLayout) Cardlayout.getLayout();
									card.show(Cardlayout, "panel_2");
									clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2,
											password_p2);

									countf = 0;
									String userid = table_userid.getValueAt(i, 2).toString();
									String mailbox = table_userid.getValueAt(i, 3).toString();
									DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + userid);
									root.add(node);
									try {
										if (mailbox.equalsIgnoreCase("Mailbox")) {
											readexchangeB_R(node, userid);

										} else if (mailbox.equalsIgnoreCase("Public Folder")) {
											readexchangePublicB_R(node, userid);
										} else {
											readexchangArchieve(node, userid);
										}
									} catch (Exception e) {

										continue;
									}

									model.reload();

									clientforexchange_input.dispose();

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
							if (countf == 0) {

								JOptionPane.showMessageDialog(mf,
										"Your user-id does not seem to have  the required permission please provide such before continuing",
										messageboxtitle, JOptionPane.ERROR_MESSAGE,
										new ImageIcon(backup.class.getResource("/information.png")));

								CardLayout card = (CardLayout) Cardlayout.getLayout();

								card.show(Cardlayout, "panel_tablecontentBR");
							}
							tree.expandRow(0);
							tree.expandAll();
							obThmailbox.close();
						}
					};

					sw1.execute();

				} else {

					JOptionPane.showMessageDialog(frame, "Please select the user id you want to convert",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(backup.class.getResource("/information.png")));
				}

			}
		});
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
		panel_tablecontentBR.add(btnNext);

		btn_previous = new JButton("");
		btn_previous.setBounds(789, 571, 142, 38);
		btn_previous.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DefaultTableModel model = (DefaultTableModel) table_userid.getModel();
				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}
				userList.clear();
				CardLayout card1 = (CardLayout) inner_cardlayout.getLayout();
				card1.show(inner_cardlayout, "panel_loginpanel");
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_1");
			}
		});
		btn_previous.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous.setIcon(new ImageIcon(backup.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
		btn_previous.setRolloverEnabled(false);
		btn_previous.setRequestFocusEnabled(false);
		btn_previous.setOpaque(false);
		btn_previous.setFocusable(false);
		btn_previous.setFocusTraversalKeysEnabled(false);
		btn_previous.setFocusPainted(false);
		btn_previous.setDefaultCapable(false);
		btn_previous.setContentAreaFilled(false);
		btn_previous.setBorderPainted(false);
		panel_tablecontentBR.add(btn_previous);

		JLabel label_7 = new JLabel("");
		label_7.setBounds(0, 552, 1075, 68);
		label_7.setIcon(new ImageIcon(backup.class.getResource("/bottom.png")));
		panel_tablecontentBR.add(label_7);
		Cursor cursor = new Cursor(Cursor.HAND_CURSOR);

		panel_9 = new JPanel();
		panel_9.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Select Backup Option",
				TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_9.setBackground(Color.WHITE);
		panel_9.setBounds(884, 74, 188, 475);
		panel_tablecontentBR.add(panel_9);
		panel_9.setLayout(null);

		rdbtnMailbox = new JRadioButton("Mailbox");
		rdbtnMailbox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_userid.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Mailbox", i, 3);
				}
			}
		});
		buttonGroup.add(rdbtnMailbox);
		rdbtnMailbox.setForeground(Color.BLACK);
		rdbtnMailbox.setFont(new Font("Arial", Font.PLAIN, 16));
		rdbtnMailbox.setBackground(Color.WHITE);
		rdbtnMailbox.setBounds(36, 39, 129, 27);
		panel_9.add(rdbtnMailbox);

		rdbtnArchieve = new JRadioButton("In-Place Archive");
		rdbtnArchieve.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_userid.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Archive Folder", i, 3);
				}
			}
		});
		buttonGroup.add(rdbtnArchieve);
		rdbtnArchieve.setForeground(Color.BLACK);
		rdbtnArchieve.setFont(new Font("Arial", Font.PLAIN, 14));
		rdbtnArchieve.setBackground(Color.WHITE);
		rdbtnArchieve.setBounds(34, 85, 129, 27);
		panel_9.add(rdbtnArchieve);

		rdbtnPublicFolder = new JRadioButton("Public Folder");
		rdbtnPublicFolder.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableModel m = table_userid.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt("Public Folder", i, 3);
				}
			}
		});
		buttonGroup.add(rdbtnPublicFolder);
		rdbtnPublicFolder.setForeground(Color.BLACK);
		rdbtnPublicFolder.setFont(new Font("Arial", Font.PLAIN, 16));
		rdbtnPublicFolder.setBackground(Color.WHITE);
		rdbtnPublicFolder.setBounds(34, 133, 129, 27);
		panel_9.add(rdbtnPublicFolder);

		JEditorPane dtrpnIfYouWant = new JEditorPane();
		dtrpnIfYouWant.setEditable(false);
		dtrpnIfYouWant.setForeground(Color.WHITE);
		dtrpnIfYouWant.setBackground(Color.DARK_GRAY);
		dtrpnIfYouWant.setFont(new Font("Arial", Font.PLAIN, 14));
		dtrpnIfYouWant.setText("If  you  want  to  take  thebackup  of Mailbox " + System.lineSeparator()
				+ ", Public Folder or " + System.lineSeparator() + "Archive Folder, then" + System.lineSeparator()
				+ "Select from given " + System.lineSeparator() + "options.");
		dtrpnIfYouWant.setBounds(10, 345, 168, 130);
		panel_9.add(dtrpnIfYouWant);

		JLabel lblNotes = new JLabel("NOTES :");
		lblNotes.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNotes.setBounds(10, 307, 78, 27);
		panel_9.add(lblNotes);

		JPanel panel_8 = new JPanel();
		panel_8.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Import CSV File",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_8.setBackground(Color.WHITE);
		panel_8.setBounds(4, 2, 883, 51);
		panel_tablecontentBR.add(panel_8);
		panel_8.setLayout(null);

		textField_3 = new JTextField();
		textField_3.setBounds(14, 16, 419, 24);
		panel_8.add(textField_3);
		textField_3.setEditable(false);
		textField_3.setColumns(10);

		JButton btnImportCsv = new JButton("");
		btnImportCsv.setBounds(441, 6, 120, 38);
		panel_8.add(btnImportCsv);
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

		JButton selectAll = new JButton("");
		selectAll.setBounds(558, 7, 122, 38);
		panel_8.add(selectAll);
		selectAll.setRequestFocusEnabled(false);
		selectAll.setRolloverEnabled(false);
		selectAll.setOpaque(false);
		selectAll.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				table_userid.convertColumnIndexToModel(0);
				TableModel m = table_userid.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt(true, i, 0);
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

		JButton unselectall = new JButton("");
		unselectall.setBounds(678, 6, 112, 38);
		panel_8.add(unselectall);
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

				table_userid.convertColumnIndexToModel(0);
				TableModel m = table_userid.getModel();

				for (int i = 0; i < m.getRowCount(); i++) {
					m.setValueAt(false, i, 0);
				}

			}
		});
		unselectall.setFocusTraversalKeysEnabled(false);
		unselectall.setFocusable(false);
		unselectall.setFocusPainted(false);
		unselectall.setDefaultCapable(false);
		unselectall.setContentAreaFilled(false);
		unselectall.setBorderPainted(false);

		JLabel lblSampleCsv = new JLabel("Sample CSV");
		lblSampleCsv.setBounds(796, 15, 85, 23);
		panel_8.add(lblSampleCsv);
		lblSampleCsv.setForeground(Color.RED);
		lblSampleCsv.setCursor(cursor);
		lblSampleCsv.setFont(new Font("Tahoma", Font.BOLD, 13));

		JLabel lblListOfAll = new JLabel("List of all your Mailbox/Office 365 Group");
		lblListOfAll.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblListOfAll.setBounds(6, 55, 264, 23);
		panel_tablecontentBR.add(lblListOfAll);

		lblNewLabel_7 = new JLabel("for Mailbox user Id greater than 100 please use CSV option\r\n");
		lblNewLabel_7.setForeground(Color.RED);
		lblNewLabel_7.setBounds(592, 52, 291, 23);
		panel_tablecontentBR.add(lblNewLabel_7);
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
		btnImportCsv.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
				chooser.setMultiSelectionEnabled(true);
				chooser.setAcceptAllFileFilterUsed(false);
				chooser.setFileFilter(new FileNameExtensionFilter("CSV File", "csv"));
				int ret = chooser.showOpenDialog(backup.this);

				Executors.newSingleThreadExecutor().execute(new Runnable() {
					public void run() {

						if (ret == JFileChooser.APPROVE_OPTION) {
							File file = new File(chooser.getSelectedFile() + "");
							if (file.exists()) {
								textField_3.setText(file.getAbsolutePath());
								readDataLineByLine(file.getAbsolutePath());
							} else {
								JOptionPane.showMessageDialog(backup.this, "  File doesn't exists!!", messageboxtitle,
										JOptionPane.ERROR_MESSAGE,
										new ImageIcon(backup.class.getResource("/information.png")));
								return;
							}
						}
					}
				});
			}
		});

		panel_2 = new JPanel();
		panel_2.setBackground(Color.WHITE);
		Cardlayout.add(panel_2, "panel_2");
		panel_2.setLayout(null);

		lbl_header = new JLabel("");
		lbl_header.setFont(new Font("Tahoma", Font.BOLD, 15));
		lbl_header.setBounds(385, 13, 368, 23);
		panel_2.add(lbl_header);

		JPanel panel_2_2 = new JPanel();
		panel_2_2.setBounds(0, 49, 1075, 571);
		panel_2.add(panel_2_2);
		panel_2_2.setLayout(new CardLayout(0, 0));

		panel_2_2_1 = new JPanel();
		panel_2_2_1.setBackground(Color.WHITE);
		panel_2_2.add(panel_2_2_1, "panel_2_2_1");
		panel_2_2_1.setLayout(null);

		label_13 = new JLabel("");
		label_13.setVisible(false);
		label_13.setIcon(new ImageIcon(backup.class.getResource("/progress-bar.gif")));
		label_13.setBounds(4, 543, 824, 22);
		panel_2_2_1.add(label_13);

		JScrollPane scrollPane_fortree_p2 = new JScrollPane();
		scrollPane_fortree_p2.setBounds(0, 0, 1065, 488);
		panel_2_2_1.add(scrollPane_fortree_p2);

		tree = new CheckboxTree();
		tree.setShowsRootHandles(true);

		scrollPane_fortree_p2.setViewportView(tree);

		tree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("root folder") {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			{
			}
		}));

		btn_next_pane2 = new JButton("");
		btn_next_pane2.setBounds(954, 526, 118, 39);
		panel_2_2_1.add(btn_next_pane2);
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
				btn_next_pane2.setIcon(new ImageIcon(backup.class.getResource("/next-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_next_pane2.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));
			}
		});

		btn_next_pane2.setIcon(new ImageIcon(backup.class.getResource("/next-btn.png")));

		btn_next_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_Previous_pane2 = new JButton("");
		btn_Previous_pane2.setBounds(838, 526, 118, 39);
		panel_2_2_1.add(btn_Previous_pane2);
		btn_Previous_pane2.setRolloverEnabled(false);
		btn_Previous_pane2.setRequestFocusEnabled(false);
		btn_Previous_pane2.setOpaque(false);
		btn_Previous_pane2.setFocusable(false);
		btn_Previous_pane2.setFocusTraversalKeysEnabled(false);
		btn_Previous_pane2.setFocusPainted(false);
		btn_Previous_pane2.setDefaultCapable(false);
		btn_Previous_pane2.setContentAreaFilled(false);
		btn_Previous_pane2.setBorderPainted(false);
		btn_Previous_pane2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Previous_pane2.setIcon(new ImageIcon(backup.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Previous_pane2.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
			}
		});

		btn_Previous_pane2.setIcon(new ImageIcon(backup.class.getResource("/previous-btn.png")));
		btn_Previous_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));

		JLabel lblNewLabel_6 = new JLabel("");
		lblNewLabel_6.setBounds(0, 503, 1247, 68);
		lblNewLabel_6.setIcon(new ImageIcon(backup.class.getResource("/bottom.png")));
		panel_2_2_1.add(lblNewLabel_6);

		btn_next_pane2.setToolTipText("Click here to Go to the Previous panel ");
		btn_Previous_pane2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				passwordField_p2.setEditable(true);
				buttonGroup.clearSelection();

				// textField.setEnabled(true);
				textField_Domainname_p2.setEnabled(true);
				textField_username_p2.setEnabled(true);
				passwordField_p2.setEnabled(true);
				textField_Domainname_p2.setText("");
				textField_username_p2.setText("");
				passwordField_p2.setText("");
				try {
					if (input) {
						if (fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("Live Exchange")
								|| fileoption.equalsIgnoreCase("Hotmail")) {
							clientforexchange_input.dispose();

						} else {
							iconnforimap_input.dispose();

						}
					}
				} catch (Exception e1) {
				}

				buttonGroup.clearSelection();

				tree.getCheckingModel();
				file = null;
				tree.clearSelection();
				DefaultTreeModel model1s = (DefaultTreeModel) tree.getModel();

				DefaultMutableTreeNode root = (DefaultMutableTreeNode) model1s.getRoot();
				root.removeAllChildren();
				model1s.reload();
				TreePath[] ac = new TreePath[0];
				tree.setCheckingPaths(ac);
				CardLayout card = (CardLayout) Cardlayout.getLayout();

				card.show(Cardlayout, "panel_tablecontentBR");

			}
		});
		btn_next_pane2.setToolTipText("Click here to Go to the Convertion ");
		btn_next_pane2.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				TreePath[] tp = tree.getCheckingPaths();

				TreePath[] checktp11 = tree.getCheckingPaths();

				if (checktp11.length == 0) {
					JOptionPane.showMessageDialog(frame, "Select File From the Tree", messageboxtitle,
							JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				} else {
					stop = false;
					hm = new HashMap<String, List<String>>();
					hmexpofinal = new HashMap<String, List<ExchangeFolderInfo>>();
					List<ExchangeFolderInfo> finallistexchangefolder = null;
					List<String> finallistfolderpath = null;
					List<ExchangeFolderInfo> listexchangetreefolder = null;
					List<String> liststringtreepath = null;
					// listExchangdinal.clear();

					for (int i = 0; i < tp.length; i++) {

						String[] str = (tp[i].toString().replace("<html><b>", "").replaceAll("[\\[\\]]", ""))
								.split(",");

						String pathoffile = "";
						String pathews = "";
						String userid = "";
						for (int j = 1; j < str.length; j++) {
							if (j == 1) {
								userid = str[j].trim();

							}

							if (hm.containsKey(userid)) {
								finallistfolderpath = hm.get(userid);
								finallistexchangefolder = hmexpofinal.get(userid);
								liststringtreepath = hmstring.get(userid);
								listexchangetreefolder = hmexpo.get(userid);

								if (j == 2) {
									pathoffile = str[j].trim();
									pathews = getRidOfIllegalFileNameCharacters(str[j].trim());
									int kki = liststringtreepath.indexOf(pathoffile);
									if (!finallistfolderpath.contains(pathoffile)) {
										finallistexchangefolder.add(listexchangetreefolder.get(kki));
										finallistfolderpath.add(pathews);
									}
								} else if (j > 2) {
									pathoffile = pathoffile + File.separator + str[j].trim();
									pathews = pathews + File.separator
											+ getRidOfIllegalFileNameCharacters(str[j].trim());
									int kki = liststringtreepath.indexOf(pathoffile);
									if (!finallistfolderpath.contains(pathoffile)) {
										try {
											finallistexchangefolder.add(listexchangetreefolder.get(kki));
											finallistfolderpath.add(pathews);
										} catch (Exception e) {
											continue;
										}
									}

								}

								hm.put(userid, finallistfolderpath);
								hmexpofinal.put(userid, finallistexchangefolder);
							} else {
								finallistfolderpath = new ArrayList<String>();
								finallistexchangefolder = new ArrayList<ExchangeFolderInfo>();
								liststringtreepath = hmstring.get(userid);
								listexchangetreefolder = hmexpo.get(userid);
								if (j == 2) {
									pathoffile = str[j].trim();
									pathews = getRidOfIllegalFileNameCharacters(str[j].trim());
									int kki = liststringtreepath.indexOf(pathoffile);
									if (!finallistfolderpath.contains(pathoffile)) {
										finallistexchangefolder.add(listexchangetreefolder.get(kki));
										finallistfolderpath.add(pathews);
									}
								} else if (j > 2) {
									pathoffile = pathoffile + File.separator + str[j].trim();
									pathews = pathews + File.separator
											+ getRidOfIllegalFileNameCharacters(str[j].trim());
									int kki = liststringtreepath.indexOf(pathoffile);
									if (!finallistfolderpath.contains(pathoffile)) {
										finallistexchangefolder.add(listexchangetreefolder.get(kki));
										finallistfolderpath.add(pathews);
									}

								}

								hm.put(userid, finallistfolderpath);
								hmexpofinal.put(userid, finallistexchangefolder);

							}

						}

					}

					listmail.clear();
					listmapi.clear();

					// listImapmesinfo.clear();
					btn_signout_p3.setVisible(false);

					System.gc();
					panel_progress.setVisible(false);

					panel_3_2.setVisible(false);
					panel_3_.setVisible(false);
					panel_3_1_1.setVisible(false);

					panel_3_.setVisible(true);

					CardLayout card1 = (CardLayout) panel_3_.getLayout();
					card1.show(panel_3_, "panel_3_1_2");
					panel_3_2.setVisible(true);
					panel_progress.setVisible(true);
					filetype = "PST";
					btn_converter.setEnabled(true);

					tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_3");
					// chckbxRemoveDuplicacy.setSelected(true);
				}

			}
		});

		panel_3 = new JPanel();
		panel_3.setBackground(Color.WHITE);
		Cardlayout.add(panel_3, "panel_3");
		panel_3.setLayout(null);

		JLabel lblSavesbackupmigrateAs = new JLabel(" Saves/Backup/Migrate As :");
		lblSavesbackupmigrateAs.setForeground(Color.BLUE);
		lblSavesbackupmigrateAs.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblSavesbackupmigrateAs.setBounds(31, 15, 212, 29);
		panel_3.add(lblSavesbackupmigrateAs);

		btn_signout_p3 = new JButton("");
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
		btn_signout_p3.setToolTipText("Click here to Sign Out of the Server ");
		btn_signout_p3.setRolloverEnabled(false);
		btn_signout_p3.setRequestFocusEnabled(false);
		btn_signout_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					String warn = "Do you want to Sign Out?";
					int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
							JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
					if (ans == JOptionPane.YES_OPTION) {

						try {
							if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
									|| filetype.equals("Hotmail")) {
								clientforexchange_output.dispose();

							} else {

								iconnforimap_output.dispose();
							}
						} catch (Exception e) {

						}

						textField_domain_name_p3.setText("");
						passwordField_p3.setText("");
						textField_domain_name_p3.setText("");
						btn_signout_p3.setVisible(false);
						btn_converter.setEnabled(false);

						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_1");
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
		btn_previous_p3 = new JButton("");
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
		btn_previous_p3.setBounds(817, 569, 118, 39);
		panel_3.add(btn_previous_p3);
		btn_previous_p3.setFont(new Font("Tahoma", Font.BOLD, 12));
		btn_previous_p3.setToolTipText("Click here to Go to the Previous Panel ");
		btn_previous_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				try {
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
				card.show(Cardlayout, "panel_2");

			}
		});
		btn_signout_p3.setBounds(912, 13, 142, 38);
		btn_signout_p3.setVisible(false);
		panel_3.add(btn_signout_p3);

		panel_3_ = new JPanel();
		panel_3_.setBounds(12, 61, 1056, 321);
		panel_3.add(panel_3_);
		panel_3_.setLayout(new CardLayout(0, 0));

		panel_3_1_1 = new JPanel();
		panel_3_1_1.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_1_1.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_1, "panel_3_1_1");
		panel_3_1_1.setLayout(null);

		JLabel lblNewLabel = new JLabel("User Name (Email Address)");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel.setBounds(34, 68, 258, 32);
		panel_3_1_1.add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel("Password ");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel_1.setBounds(34, 98, 278, 32);
		panel_3_1_1.add(lblNewLabel_1);

		panel_3_1_1_1 = new JPanel();
		panel_3_1_1_1.setBackground(Color.WHITE);
		panel_3_1_1_1.setBounds(12, 11, 755, 46);
		panel_3_1_1.add(panel_3_1_1_1);
		panel_3_1_1_1.setLayout(null);

		lbl_Domain = new JLabel("");
		lbl_Domain.setFont(new Font("Tahoma", Font.BOLD, 14));
		lbl_Domain.setBounds(10, 11, 276, 26);
		panel_3_1_1_1.add(lbl_Domain);

		textField_domain_name_p3 = new JTextField();
		textField_domain_name_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_domain_name_p3.setComponentPopupMenu(menu);
		textField_domain_name_p3.setBounds(298, 11, 424, 26);
		panel_3_1_1_1.add(textField_domain_name_p3);
		textField_domain_name_p3.setColumns(10);

		textField_username_p3 = new JTextField();
		textField_username_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_username_p3.setComponentPopupMenu(menu);
		textField_username_p3.setBounds(313, 68, 423, 25);
		panel_3_1_1.add(textField_username_p3);
		textField_username_p3.setColumns(10);

		passwordField_p3 = new JPasswordField();
		passwordField_p3.setHorizontalAlignment(JTextField.CENTER);
		passwordField_p3.setComponentPopupMenu(menu);
		passwordField_p3.setBounds(313, 104, 423, 25);
		panel_3_1_1.add(passwordField_p3);

		chckbxShowPassword_p3 = new JCheckBox("Show Password");
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
		chckbxShowPassword_p3.setBounds(745, 107, 143, 25);
		panel_3_1_1.add(chckbxShowPassword_p3);

		btn_Sign_p3 = new JButton("");
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

		btn_Sign_p3.setBounds(778, 196, 118, 39);
		panel_3_1_1.add(btn_Sign_p3);
		btn_Sign_p3.setToolTipText("Click here to Sign In the Server ");
		btn_Sign_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				username_p3 = "";
				password_p3 = "";
				domain_p3 = "";
				comboBox_fileDestination_type.setEnabled(false);
				textField_domain_name_p3.setEnabled(false);
				textField_username_p3.setEnabled(false);
				passwordField_p3.setEnabled(false);
				tf_portNo_p3.setEnabled(false);
				chckbxShowPassword_p3.setEnabled(false);
				btn_Sign_p3.setEnabled(false);
				try {
					domain_p3 = textField_domain_name_p3.getText().replaceAll("//s", "");
					// System.out.println(domain_p3);
				} catch (Exception a) {

				}
				try {
					username_p3 = textField_username_p3.getText().replaceAll("//s", "");

				} catch (Exception a) {

				}
				try {
					password_p3 = new String(passwordField_p3.getPassword());

				} catch (Exception a) {

				}
				try {
					portnofiletype = Integer.parseInt(tf_portNo_p3.getText().replaceAll("//s", ""));
				} catch (Exception a) {

				}
				if (username_p3.equalsIgnoreCase("") || password_p3.equalsIgnoreCase("")) {

					if (username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(frame, "User name and Password fields cannot be empty",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (username_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "User name field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (password_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "Password field cannot be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					}

					comboBox_fileDestination_type.setEnabled(true);
					chckbxShowPassword_p3.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					textField_domain_name_p3.setEnabled(true);
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					tf_portNo_p3.setEnabled(true);
					btn_converter.setEnabled(false);
					btn_Sign_p3.setEnabled(true);
					btn_signout_p3.setVisible(false);
					chckbx_Mail_Filter.setEnabled(true);
					task_box.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("Live Exchange") && domain_p3.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "Computer Name or IP Address field can not be empty",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_fileDestination_type.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("IMAP") && domain_p3.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "IMAP Host field can not be empty", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_fileDestination_type.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("IMAP") && tf_portNo_p3.getText().isEmpty()) {

					JOptionPane.showMessageDialog(frame, "Port No field cann't be empty", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_fileDestination_type.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (!(isValid(username_p3))) {

					JOptionPane.showMessageDialog(frame, "Please enter a valid username", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_fileDestination_type.setEnabled(true);
					chckbxShowPassword_p3.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					textField_domain_name_p3.setEnabled(true);
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					tf_portNo_p3.setEnabled(true);
					btn_Sign_p3.setEnabled(true);

				} else {

					th = new Thread(new Runnable() {

						@Override
						public void run() {
							lbl_connecting_p3.setVisible(true);

							try {
								btn_converter.setEnabled(false);
								btn_previous_p3.setEnabled(false);

								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									conntiontooffice365_output();
								}

								else if (filetype.equalsIgnoreCase("GMAIL")) {

									connectiontogmail_output();

								} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
									connectiontoYandex_output();

								} else if (filetype.equalsIgnoreCase("Hostgator email")) {
									connectiontoHostgator_output();

								} else if (filetype.equalsIgnoreCase("Icloud")) {
									connectiontoicloud_output();

								} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
									connectiontoGoDaddy_output();

								} else if (filetype.equalsIgnoreCase("Live Exchange")) {
									connectionwithexchangeserver_output();

								} else if (filetype.equalsIgnoreCase("IMAP")) {

									connectiontoimap_output();

								} else if (filetype.equalsIgnoreCase("Hotmail")) {
									conntiontohotmail_output();

								} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

									connectiontozoho_output();

								} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

									connectiontoyahoo_output();

								} else if (filetype.equalsIgnoreCase("AOL")) {

									connectiontoaol_output();

								} else if (filetype.equalsIgnoreCase("Amazon Webmail")) {

									connectiontoinaws_output();

								}

								panel_3_1_1.setVisible(false);
								panel_3_1_2.setVisible(true);
								btn_converter.setEnabled(true);
								btn_signout_p3.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
							} catch (Error e) {
								mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
								lblLive_Chat_p3.setVisible(true);
								btn_signout_p3.setVisible(false);
								btn_signout_p3.setEnabled(false);
								// System.out.println(e.getMessage());
								comboBox_fileDestination_type.setEnabled(true);

								if (filetype.equalsIgnoreCase("Gmail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Gmail please check your Credantial OR Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_3_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(frame, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(frame, "Connection not established", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}
								// System.out.println(e.getMessage());
								e.printStackTrace();

							} finally {
								lbl_connecting_p3.setVisible(false);

								chckbxShowPassword_p3.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								textField_domain_name_p3.setEnabled(true);
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								tf_portNo_p3.setEnabled(true);
								btn_Sign_p3.setEnabled(true);

							}

						}
					});
					th.start();

				}
			}
		});
		btn_Sign_p3.setFont(new Font("Tahoma", Font.BOLD, 14));

		lbl_connecting_p3 = new JLabel("");
		lbl_connecting_p3.setBounds(471, 197, 85, 32);
		panel_3_1_1.add(lbl_connecting_p3);
		lbl_connecting_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));

		tf_portNo_p3 = new JTextField();
		tf_portNo_p3.setHorizontalAlignment(JTextField.CENTER);
		tf_portNo_p3.setComponentPopupMenu(menu);
		tf_portNo_p3.setBounds(313, 147, 423, 25);
		tf_portNo_p3.setText(Integer.toString(993));
		panel_3_1_1.add(tf_portNo_p3);
		tf_portNo_p3.setColumns(10);

		lblPortNo = new JLabel("Port No");
		lblPortNo.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblPortNo.setBounds(34, 151, 73, 27);
		panel_3_1_1.add(lblPortNo);

		lblLive_Chat_p3 = new JLabel("More Help");
		lblLive_Chat_p3.setForeground(Color.RED);
		lblLive_Chat_p3.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblLive_Chat_p3.setBounds(844, 11, 65, 25);
		lblLive_Chat_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");
			}
		});

		panel_3_1_1.add(lblLive_Chat_p3);

		panel = new JPanel();
		panel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBackground(Color.WHITE);
		panel.setBounds(31, 258, 995, 52);
		panel_3_1_1.add(panel);
		panel.setLayout(null);

		lblMakeSureYou = new JLabel("Please  Click on The Link");
		lblMakeSureYou.setForeground(Color.BLACK);
		lblMakeSureYou.setBounds(10, 11, 174, 32);
		panel.add(lblMakeSureYou);
		lblMakeSureYou.setFont(new Font("Tahoma", Font.BOLD, 13));

		lblEnableImap_p3 = new JLabel("To Enable IMAP");
		lblEnableImap_p3.setBounds(820, 17, 107, 25);
		panel.add(lblEnableImap_p3);
		lblEnableImap_p3.setForeground(Color.BLUE);
		lblEnableImap_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo_p3 = new JLabel("Turn Off Two Step Verification");
		lblTurnOffTwo_p3.setBounds(205, 11, 533, 32);
		panel.add(lblTurnOffTwo_p3);
		lblTurnOffTwo_p3.setCursor(mf.cursor);
		lblTurnOffTwo_p3.setForeground(Color.BLUE);
		lblTurnOffTwo_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				mf.openbrowserturntwostepoff(filetype);
			}
		});
		lblTurnOffTwo_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));
		lblEnableImap_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				mf.openbrowserenableimap(filetype);

			}
		});
		lbl_connecting_p3.setVisible(false);

		panel_3_1_2 = new JPanel();
		panel_3_1_2.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_1_2.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_2, "panel_3_1_2");
		panel_3_1_2.setLayout(null);

		panel_3_2 = new JPanel();
		panel_3_2.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_3_2.setBackground(Color.WHITE);
		panel_3_2.setBounds(12, 395, 1058, 48);
		panel_3.add(panel_3_2, "panel_3_2");
		panel_3_2.setLayout(null);
		panel_3_2.setVisible(false);

		panel_mailfilter = new JPanel();
		panel_mailfilter.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_mailfilter.setBackground(Color.WHITE);
		panel_mailfilter.setBounds(494, 34, 556, 39);
		panel_3_1_2.add(panel_mailfilter);
		panel_mailfilter.setLayout(null);

		JLabel label_1 = new JLabel("Start Date");
		label_1.setFont(new Font("Tahoma", Font.PLAIN, 11));
		label_1.setBounds(10, 11, 93, 20);
		panel_mailfilter.add(label_1);

		JLabel label_3 = new JLabel("End Date\r\n");
		label_3.setFont(new Font("Tahoma", Font.PLAIN, 11));
		label_3.setBounds(318, 7, 102, 19);
		panel_mailfilter.add(label_3);

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
		dateChooser_mail_fromdate.setBackground(Color.WHITE);
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
		dateChooser_mail_fromdate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_fromdate.setBounds(76, 9, 176, 23);
		dateChooser_mail_fromdate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_fromdate);

		dateChooser_mail_tilldate = new JDateChooser();
		dateChooser_mail_tilldate.getCalendarButton().setBackground(Color.WHITE);
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
				} catch (Error e1) {
					mf.logger.warning("Error :" + e1.getMessage() + System.lineSeparator());
				} catch (Exception e1) {
					mf.logger.warning("Exception :" + e1.getMessage() + System.lineSeparator());
					return;
				}
			}
		});
		dateChooser_mail_tilldate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_tilldate.setBounds(393, 10, 149, 19);
		dateChooser_mail_tilldate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_tilldate);

		panel_Calender = new JPanel();
		panel_Calender.setBackground(Color.WHITE);
		panel_Calender.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Calender Filter",
				TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_Calender.setBounds(34, 108, -12, 18);
		panel_Calender.setEnabled(false);
		panel_Calender.setVisible(false);
		panel_3_1_2.add(panel_Calender);
		panel_Calender.setLayout(null);

		JDateChooser dateChooser_calender_start = new JDateChooser();
		dateChooser_calender_start.setBackground(Color.WHITE);
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
		dateChooser_calender_start.setBounds(169, 19, 168, 20);
		panel_Calender.add(dateChooser_calender_start);
		dateChooser_calender_start.setEnabled(false);
		dateChooser_calender_start.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));

		JDateChooser dateChooser_calendar_end = new JDateChooser();
		dateChooser_calendar_end.setBackground(Color.WHITE);
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
					Calendar calendarstartdate = dateChooser_calender_start.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_calendar_end.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					mf.logger.warning("Error :" + e1.getMessage() + System.lineSeparator());
				}

				catch (Exception e1) {
					mf.logger.warning("Exceptin :" + e1.getMessage() + System.lineSeparator());
					return;
				}
			}
		});
		dateChooser_calendar_end.setBounds(808, 28, 149, 19);
		panel_Calender.add(dateChooser_calendar_end);
		dateChooser_calendar_end.setEnabled(false);
		dateChooser_calendar_end.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));

		JLabel label = new JLabel("End Date\r\n");
		label.setFont(new Font("Tahoma", Font.BOLD, 15));
		label.setBounds(695, 28, 102, 19);
		panel_Calender.add(label);

		JLabel label_2 = new JLabel("Start Date");
		label_2.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_2.setBounds(76, 19, 114, 20);
		panel_Calender.add(label_2);

		JPanel panel_7 = new JPanel();
		panel_7.setBackground(Color.WHITE);
		panel_7.setBounds(39, 94, 1, 11);
		panel_7.setVisible(false);
		panel_3_1_2.add(panel_7);
		panel_7.setLayout(null);

		chckbx_calender_box = new JCheckBox("Calender Filter");
		chckbx_calender_box.setBackground(Color.WHITE);
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
		chckbx_calender_box.setBounds(8, 0, 0, 25);
		panel_7.add(chckbx_calender_box);

		task_box = new JCheckBox("Task Filter");
		task_box.setFont(new Font("Tahoma", Font.BOLD, 12));
		task_box.setBounds(497, 85, 122, 18);
		panel_3_1_2.add(task_box);
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

		panel_taskfilter = new JPanel();
		panel_taskfilter.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_taskfilter.setBackground(Color.WHITE);
		panel_taskfilter.setBounds(497, 108, 552, 39);
		panel_3_1_2.add(panel_taskfilter);
		panel_taskfilter.setLayout(null);

		dateChooser_task_start_date = new JDateChooser();
		dateChooser_task_start_date.setBounds(83, 11, 169, 22);
		panel_taskfilter.add(dateChooser_task_start_date);
		dateChooser_task_start_date.setBackground(Color.WHITE);
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

		JLabel label_4 = new JLabel("Start Date");
		label_4.setBounds(10, 11, 114, 20);
		panel_taskfilter.add(label_4);
		label_4.setFont(new Font("Tahoma", Font.PLAIN, 11));

		JLabel label_5 = new JLabel("End Date\r\n");
		label_5.setBounds(321, 14, 58, 19);
		panel_taskfilter.add(label_5);
		label_5.setFont(new Font("Tahoma", Font.PLAIN, 11));

		dateChooser_task_end_date = new JDateChooser();
		dateChooser_task_end_date.setBounds(391, 13, 152, 22);
		panel_taskfilter.add(dateChooser_task_end_date);
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
		dateChooser_task_end_date.setBackground(Color.WHITE);
		dateChooser_task_end_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_task_end_date.setMaxSelectableDate(enddate);

				try {
					Calendar calendarstartdate = dateChooser_task_start_date.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_task_end_date.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					mf.logger.warning("Error :" + e1.getMessage() + System.lineSeparator());
				} catch (Exception e1) {
					mf.logger.warning("Exception :" + e1.getMessage() + System.lineSeparator());
					return;
				}

			}
		});
		dateChooser_task_end_date.setEnabled(false);

		chckbx_Mail_Filter = new JCheckBox("Mail Filter");
		chckbx_Mail_Filter.setFont(new Font("Tahoma", Font.BOLD, 12));
		chckbx_Mail_Filter.setBounds(494, 5, 113, 25);
		panel_3_1_2.add(chckbx_Mail_Filter);
		chckbx_Mail_Filter.setBackground(Color.WHITE);

		JPanel panel_5 = new JPanel();
		panel_5.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_5.setBackground(Color.WHITE);
		panel_5.setVisible(false);
		panel_5.setBounds(494, 224, 556, 48);
		panel_3_1_2.add(panel_5);
		panel_5.setLayout(null);

		comboBox = new JComboBox<String>();
		comboBox.setBounds(187, 15, 274, 20);
		panel_5.add(comboBox);
		comboBox.addItem("Subject");
		comboBox.addItem("Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_Subject");
		comboBox.addItem("From_Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("From_Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("From_Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("From_Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_From_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_From_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_From_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_From_Subject");
		comboBox.setVisible(false);
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

		comboBox_fileDestination_type = new JComboBox(Main_Frame.sfd);
		mf.imageMap_output = mf.createImageMap_output(mf.l_output);
		comboBox_fileDestination_type.setRenderer(mf.new ListRenderer_output());
		comboBox_fileDestination_type.setBackground(Color.WHITE);
		comboBox_fileDestination_type.setFont(new Font("Tahoma", Font.BOLD, 15));
		comboBox_fileDestination_type.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				panel_3_.setVisible(false);

				panel_3_2.setVisible(false);

				panel_3_1_2.setVisible(false);

				panel_3_1_1.setVisible(false);

				panel_3_1_1_1.setVisible(false);

				btnStop.setVisible(false);

				lblPortNo.setVisible(false);

				tf_portNo_p3.setVisible(false);

				textField_domain_name_p3.setText("");
				panel_5.setVisible(false);
				textField_username_p3.setText("");
				panel_progress.setVisible(true);
				lblMakeSureYou.setVisible(true);
				lblEnableImap_p3.setVisible(true);
				lblTurnOffTwo_p3.setVisible(true);

				tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
				passwordField_p3.setText("");
				comboBox.setVisible(false);
				btn_signout_p3.setVisible(false);
				btn_converter.setEnabled(false);
				lbl_progressreport.setText("");

				if (e.getSource() == comboBox_fileDestination_type) {

					JComboBox cb = (JComboBox) e.getSource();

					filetype = (String) cb.getSelectedItem();

				}
				//
				if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
						|| filetype.equalsIgnoreCase("Amazon WorkMail") || filetype.equalsIgnoreCase("Icloud")
						|| filetype.equalsIgnoreCase("Hostgator email") || filetype.equalsIgnoreCase("GoDaddy email")
						|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("AOL")
						|| filetype.equalsIgnoreCase("Live Exchange") || filetype.equalsIgnoreCase("HOTMAIL")
						|| filetype.equalsIgnoreCase("IMAP") || filetype.equalsIgnoreCase("Zoho Mail")
						|| filetype.equalsIgnoreCase("Yandex Mail")) {
					output = true;
					btn_converter.setEnabled(false);
					lblEnableImap_p3.setVisible(false);

					lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
							+ " account , you'll need to generate and use an app password.</U></HTML>");
					lblNewLabel_1.setText("Password(Third Party App Password)");

					if (filetype.equalsIgnoreCase("OFFICE 365")) {

						lblNewLabel_1.setText("Password/Third Party App Password");

						lblTurnOffTwo_p3.setText(
								"If your multi-factor authentication is enabled, then you need to create third-party app password.");

					}
					panel.setVisible(true);
					if (filetype.equalsIgnoreCase("GoDaddy email")) {
						panel.setVisible(false);
					}
					if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("Zoho Mail")) {
						lblEnableImap_p3.setVisible(true);
						lblNewLabel_1.setText("Password(Third Party App Password)");

						lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
								+ " account , you'll need to generate and use an app password or turn on less secure app</U></HTML>");

					}
					if (filetype.equalsIgnoreCase("Live Exchange")) {
						panel_3_.setVisible(true);

						panel_3_1_1.setVisible(true);

						panel_3_1_1_1.setVisible(true);
						lblNewLabel_1.setText("Password");
						lbl_Domain.setText("  " + "IP or Computer Name");
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);

					} else if (filetype.equalsIgnoreCase("IMAP") || filetype.equalsIgnoreCase("Amazon WorkMail")) {
						panel_3_.setVisible(true);
						lblNewLabel_1.setText("Password");
						panel_3_1_1.setVisible(true);

						panel_3_1_1_1.setVisible(true);

						lbl_Domain.setText("  " + filetype + " HOST");

						lblPortNo.setVisible(true);

						tf_portNo_p3.setVisible(true);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
					} else if (filetype.equalsIgnoreCase("Hostgator email")) {
						panel_3_.setVisible(true);
						lblNewLabel_1.setText("Password");
						panel_3_1_1.setVisible(true);

						panel_3_1_1_1.setVisible(true);

						lbl_Domain.setText("  Hostgator HOST");

						lblPortNo.setVisible(true);

						tf_portNo_p3.setVisible(true);
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
					}

					else {

						panel_3_.setVisible(true);

						panel_3_1_1.setVisible(true);

					}
				} else {
					if (filetype.equalsIgnoreCase("Opera Mail")) {

						String OS = System.getProperty("os.name").toLowerCase();
						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA").replace("Roaming", "Local") + File.separator + "Opera Mail"
									+ File.separator + "Opera Mail" + File.separator + "Mail" + File.separator
									+ "store";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Application Support" + File.separator + "Opera Mail" + File.separator + "mail";
						}

						System.out.println(str);

						if (new File(str).exists()) {

							thunderbirdpath = str;
							tf_Destination_Location.setText(str);

						} else {
							String warn = filetype + " Not Installed Do you want to proced ?";
							int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.YES_OPTION) {

							} else {
								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					} else if (filetype.equalsIgnoreCase("Thunderbird")) {

						String OS = System.getProperty("os.name").toLowerCase();
						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA") + File.separator + "Thunderbird" + File.separator
									+ "Profiles";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Thunderbird" + File.separator + "Profiles";
						}

						System.out.println(str);

						if (new File(str).exists()) {

							File[] f = new File(str).listFiles();
							for (File fl : f) {
								if (fl != null) {
									if (fl.isDirectory()) {
										String filename = fl.getName();
										String extension = filename.substring(filename.lastIndexOf(".") + 1,
												filename.length());
										String ext = "default";
										if (ext.equals(extension)) {
											// System.out.println(file);

											String defaultfolder = fl.getName();

											str = str + File.separator + defaultfolder + File.separator + "Mail"
													+ File.separator + "Local Folders";
											thunderbirdpath = str;
											tf_Destination_Location.setText(str);
											break;
										} else {

										}
									}
								}
							}
						} else {
							String warn = filetype + " Not Installed Do you want to proced ?";
							int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.YES_OPTION) {

							} else {
								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					}
					output = false;
					panel_3_.setVisible(true);

					CardLayout card = (CardLayout) panel_3_.getLayout();
					card.show(panel_3_, "panel_3_1_2");

					panel_3_2.setVisible(true);
					btn_converter.setEnabled(true);

					if (!(filetype.equalsIgnoreCase("PST") || filetype.equalsIgnoreCase("Thunderbird")
							|| filetype.equalsIgnoreCase("Opera Mail") || filetype.equalsIgnoreCase("OST")
							|| filetype.equalsIgnoreCase("MBOX") || filetype.equalsIgnoreCase("CSV"))) {
						panel_5.setVisible(true);
						comboBox.setVisible(true);
					}
				}

			}
		});
		comboBox_fileDestination_type.setBounds(252, 13, 638, 35);

		panel_3.add(comboBox_fileDestination_type);

		tf_Destination_Location = new JTextField();
		tf_Destination_Location.setBackground(Color.WHITE);
		tf_Destination_Location.setBounds(198, 11, 638, 20);
		panel_3_2.add(tf_Destination_Location);
		tf_Destination_Location.setEditable(false);
		tf_Destination_Location.setColumns(10);

		btn_Destination = new JButton("");
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
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
			}
		});

		btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));

		btn_Destination.setBounds(855, 0, 162, 38);
		panel_3_2.add(btn_Destination);
		btn_Destination.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_converter = new JButton("");
		btn_converter.setRolloverEnabled(false);
		btn_converter.setEnabled(false);
		btn_converter.setRequestFocusEnabled(false);
		btn_converter.setOpaque(false);
		btn_converter.setFocusTraversalKeysEnabled(false);
		btn_converter.setFocusable(false);
		btn_converter.setFocusPainted(false);
		btn_converter.setContentAreaFilled(false);
		btn_converter.setBorderPainted(false);
		btn_converter.setDefaultCapable(false);
		btn_converter.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
			}
		});

		btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
		btn_converter.setBounds(948, 569, 118, 39);
		panel_3.add(btn_converter);
		btn_converter.setFont(new Font("Tahoma", Font.BOLD, 12));

		panel_progress = new JPanel();
		panel_progress.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_progress.setBackground(Color.WHITE);
		panel_progress.setBounds(12, 454, 1057, 87);
		panel_3.add(panel_progress);
		panel_progress.setLayout(null);

		btnStop = new JButton("");
		btnStop.setBounds(899, 11, 118, 28);
		panel_progress.add(btnStop);
		btnStop.setContentAreaFilled(false);
		btnStop.setBorderPainted(false);
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
		btnStop.setRequestFocusEnabled(false);
		btnStop.setOpaque(false);
		btnStop.setFocusable(false);
		btnStop.setFocusTraversalKeysEnabled(false);
		btnStop.setFocusPainted(false);
		btnStop.setDefaultCapable(false);

		Progressbar = new JLabel("");
		Progressbar.setBounds(10, 15, 892, 24);
		panel_progress.add(Progressbar);
		Progressbar.setVisible(false);
		Progressbar.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));
		JLabel lblnamingconvention = new JLabel("Naming Convention");
		lblnamingconvention.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblnamingconvention.setBounds(20, 14, 140, 23);
		panel_5.add(lblnamingconvention);
		SwingUtilities.invokeLater(new Runnable() {

			public void run() {

				comboBox.setSelectedItem("Subject");

				comboBox_fileDestination_type.setSelectedItem("PST");
			}
		});

		lbl_progressreport = new JLabel("");
		lbl_progressreport.setBounds(20, 50, 882, 24);
		panel_progress.add(lbl_progressreport);

		JPanel panel_duplicacy = new JPanel();
		panel_duplicacy.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_duplicacy.setBackground(Color.WHITE);
		panel_duplicacy.setBounds(4, 8, 480, 274);
		panel_3_1_2.add(panel_duplicacy);
		panel_duplicacy.setLayout(null);

		chckbxRemoveDuplicacy = new JCheckBox("Remove Duplicate Mail On basis of To, From, Subject, Bcc, Body");
		chckbxRemoveDuplicacy.setForeground(Color.RED);
		chckbxRemoveDuplicacy.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxRemoveDuplicacy.setBackground(Color.WHITE);
		chckbxRemoveDuplicacy.setBounds(6, 7, 396, 34);
		panel_duplicacy.add(chckbxRemoveDuplicacy);

		label_15 = new JLabel("");
		label_15.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_15.setToolTipText("Remove Duplicate Mail On basis of To, From, Subject, Bcc, Body");
		label_15.setBounds(424, 17, 26, 23);
		panel_duplicacy.add(label_15);

		JPanel panel_6 = new JPanel();
		panel_6.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_6.setBackground(Color.WHITE);
		panel_6.setBounds(494, 171, 557, 44);
		panel_3_1_2.add(panel_6);
		panel_6.setLayout(null);

		JLabel lblNewLabel_10 = new JLabel("");
		lblNewLabel_10.setToolTipText("Add the name by which the folder will be created " + "\r\n"
				+ " it must not contain these characters :\\?/|*<>\t");
		lblNewLabel_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lblNewLabel_10.setBounds(451, 10, 24, 25);
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
		textField_customfolder.setBounds(181, 10, 240, 23);
		panel_6.add(textField_customfolder);
		textField_customfolder.setEditable(false);
		textField_customfolder.setColumns(10);

		chckbxCustomFolderName = new JCheckBox("Custom Folder Name");
		chckbxCustomFolderName.setBounds(16, 11, 141, 23);
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

		label_12 = new JLabel("");
		label_12.setBounds(936, 48, 81, 28);
		panel_progress.add(label_12);

		label_9 = new JLabel("");
		label_9.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_9.setBounds(0, 545, 1075, 75);
		panel_3.add(label_9);
		btnStop.setToolTipText("Click here to Stop the Conversion ");
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(mf, warn, messageboxtitle, JOptionPane.YES_NO_OPTION,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					stop = true;
				}
			}
		});
		btn_converter.setToolTipText("Click here to Begin the Conversion ");
		btn_converter.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
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
					} catch (Exception e) {
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
						btn_Destination.setEnabled(false);
						btn_previous_p3.setEnabled(false);
						btn_converter.setEnabled(false);
						btnStop.setVisible(true);
						panel_5.setVisible(false);
						textField_customfolder.setEditable(false);
						chckbxCustomFolderName.setEnabled(false);
						dateChooser_calender_start.setEnabled(false);
						dateChooser_calendar_end.setEnabled(false);
						dateChooser_mail_fromdate.setEnabled(false);
						dateChooser_mail_tilldate.setEnabled(false);
						comboBox.setVisible(false);
						lbl_progressreport.setText("");
						task_box.setEnabled(false);
						chckbxRemoveDuplicacy.setEnabled(false);
						chckbx_Mail_Filter.setEnabled(false);
						dateChooser_task_start_date.setEnabled(false);
						dateChooser_task_end_date.setEnabled(false);
						chckbx_Mail_Filter.setEnabled(false);
						chckbx_calender_box.setEnabled(false);
						task_box.setEnabled(false);
						chckbx_calender_box.setEnabled(false);
						btn_signout_p3.setVisible(false);
						comboBox_fileDestination_type.setEnabled(false);
						long begintime = System.currentTimeMillis();
						Main_Frame.comboBox.setSelectedIndex(comboBox.getSelectedIndex());
						String destinationfile = "";
						String fa = "";
						Progressbar.setVisible(true);
						long totalcount = 0;
						if (fileoption.equalsIgnoreCase("MBOX")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".mbox", "").replace(".mbx", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);

						} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
							String fnam = file.getName();
							fnam = getRidOfIllegalFileNameCharacters(fnam);
							fnam = fnam.replace(".pst", "");
							fnam = fnam.replaceAll("\\p{C}", "_");

							fname = fnam.trim();

						} else if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".ost", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("EML Files (.eml)")) {
							String fnam = file.getName();
							fname = fnam.replace(".eml", "").replace("//s", "");
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("EMLX Files (.emlx)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".emlx", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("Message Files (.msg)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".msg", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("OLM files (.olm)")) {
							String fnam = file.getName();
							fnam = fnam.replace(".olm", "");
							fnam = fnam.replaceAll("\\p{C}", "_");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						}

						if (checkconvertagain) {

							if (fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("Yahoo Mail")
									|| fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("IMAP")
									|| fileoption.equalsIgnoreCase("Opera Mail")
									|| filetype.equalsIgnoreCase("Zoho Mail")
									|| fileoption.equalsIgnoreCase("Amazon Webmail")
									|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("Icloud")
									|| filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| fileoption.equalsIgnoreCase("Thunderbird") || fileoption.equalsIgnoreCase("AOL")
									|| fileoption.equalsIgnoreCase("Live Exchange")
									|| fileoption.equalsIgnoreCase("Hotmail")) {

								if (filetype.equalsIgnoreCase("GoDaddy email")) {
									calendertime = calendertime.replaceAll("[^a-zA-Z0-9]", "");
									fname = fname.replaceAll("[^a-zA-Z0-9]", "");

								}

								if (chckbxCustomFolderName.isSelected()) {

									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
									if (filetype.equalsIgnoreCase("GoDaddy email")) {
										customerfolder = customerfolder.replaceAll("[^a-zA-Z0-9]", "");

									}
									String servername = getRidOfIllegalFileNameCharacters(
											textField_username_p2.getText());
									f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
											+ File.separator + servername + filetype);

									if (!f.isFile()) {

										f.mkdirs();

										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									} else {
										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ "(" + calendertime + ")" + File.separator + servername + filetype
												+ "_" + comboBox.getSelectedItem().toString());

										f.mkdirs();

										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									}

								} else {

									f = new File(tf_Destination_Location.getText() + File.separator + calendertime
											+ File.separator + filetype + "_" + comboBox.getSelectedItem().toString());

									f.mkdirs();
									destination_path = f.getAbsolutePath();

								}

							} else {
								if (chckbxCustomFolderName.isSelected()) {

									String customerfolder = textField_customfolder.getText().replace("//s", "");

									customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
									f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
											+ filetype);

									if (!f.isFile()) {

										f.mkdirs();

										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									} else {
										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ "(" + calendertime + ")");

										f.mkdirs();

										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									}

								} else {
									calendertime = getRidOfIllegalFileNameCharacters(calendertime);
									f = new File(tf_Destination_Location.getText() + File.separator + calendertime
											+ File.separator + username_p2 + filetype);
									f.mkdirs();
									destination_path = f.getAbsolutePath();

									System.out.println(destination_path);
									destinationfile = f.getAbsolutePath();

								}
							}

						} else {
							if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
									|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("Live Exchange")
									|| fileoption.equalsIgnoreCase("Opera Mail")
									|| filetype.equalsIgnoreCase("Zoho Mail") || filetype.equalsIgnoreCase("Icloud")
									|| filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| filetype.equalsIgnoreCase("Amazon Webmail")
									|| filetype.equalsIgnoreCase("Yandex Mail")
									|| fileoption.equalsIgnoreCase("Thunderbird")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("IMAP")
									|| filetype.equalsIgnoreCase("HOTMAIL"))) {

								if (fileoption.equalsIgnoreCase("OFFICE 365 Backup & Restore")) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(
												tf_Destination_Location.getText() + File.separator + customerfolder);

										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")");

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {

										f = new File(tf_Destination_Location.getText() + File.separator + calendertime);
										f.mkdirs();
										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									}

								} else {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										f = new File(
												tf_Destination_Location.getText() + File.separator + customerfolder);

										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")");

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {
										fname = username_p2;
										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname);
										f.mkdirs();
										destination_path = tf_Destination_Location.getText() + File.separator
												+ calendertime + File.separator + fname;

										System.out.println(destination_path);
										destinationfile = tf_Destination_Location.getText() + File.separator
												+ calendertime + File.separator + fname;

									}

								}

							} else {
								if (filetype.equalsIgnoreCase("gmail") || filetype.equalsIgnoreCase("YAHOO MAIL")
										|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("Zoho Mail")
										|| filetype.equalsIgnoreCase("Amazon Webmail")
										|| filetype.equalsIgnoreCase("Icloud")
										|| filetype.equalsIgnoreCase("GoDaddy email")

										|| filetype.equalsIgnoreCase("Yandex Mail")) {

									if (chckbxCustomFolderName.isSelected()) {
										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										if (filetype.equalsIgnoreCase("GoDaddy email")) {
											customerfolder = customerfolder.replaceAll("[^a-zA-Z0-9]", "");

										}
										try {
											try {
												clientforimap_output.createFolder(iconnforimap_output, customerfolder);
											} catch (Exception e2) {
												connectionHandle1("");
												clientforimap_output.createFolder(iconnforimap_output, customerfolder);
											}

											clientforimap_output.selectFolder(iconnforimap_output, customerfolder);
											path = customerfolder;

										} catch (Exception e) {

											clientforimap_output.createFolder(iconnforimap_output,
													customerfolder + "(" + calendertime + ")");

											clientforimap_output.selectFolder(iconnforimap_output,
													customerfolder + "(" + calendertime + ")");
											path = customerfolder + "(" + calendertime + ")";

										}
									} else {

										calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
										try {
											clientforimap_output.createFolder(iconnforimap_output, calendertime);
										} catch (Exception e2) {
											connectionHandle1("");
											clientforimap_output.createFolder(iconnforimap_output, calendertime);
										}

										clientforimap_output.selectFolder(iconnforimap_output, calendertime);

										path = calendertime;

									}

								} else if (filetype.equalsIgnoreCase("IMAP")
										|| filetype.equalsIgnoreCase("Hostgator email")) {

									if (chckbxCustomFolderName.isSelected()) {
										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										try {

											try {
												clientforimap_output.createFolder(iconnforimap_output,
														"INBOX" + "." + customerfolder);
											} catch (Exception e2) {
												connectionHandle1("");
												clientforimap_output.createFolder(iconnforimap_output,
														"INBOX" + "." + customerfolder);
											}

											clientforimap_output.selectFolder(iconnforimap_output,
													"INBOX" + "." + customerfolder);
											path = "INBOX" + "." + customerfolder;

										} catch (Exception e) {

											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + "." + customerfolder + "(" + calendertime + ")");

											clientforimap_output.selectFolder(iconnforimap_output,
													"INBOX" + "." + customerfolder + "(" + calendertime + ")");
											path = "INBOX" + "." + customerfolder + "(" + calendertime + ")";

										}
									} else {

										calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

										try {
											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + "." + calendertime);
										} catch (Exception e2) {
											connectionHandle1("");
											clientforimap_output.createFolder(iconnforimap_output,
													"INBOX" + "." + calendertime);
										}

										clientforimap_output.selectFolder(iconnforimap_output,
												"INBOX" + "." + calendertime);

										path = "INBOX" + "." + calendertime;

									}

								} else if (filetype.equalsIgnoreCase("OFFICE 365")
										|| filetype.equalsIgnoreCase("Live Exchange")
										|| filetype.equalsIgnoreCase("Hotmail")) {

									if (chckbxCustomFolderName.isSelected()) {
										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										try {

											Folderuri = clientforexchange_output.createFolder(customerfolder).getUri();
											fa = Folderuri;

										} catch (Exception e) {

											Folderuri = clientforexchange_output
													.createFolder(customerfolder + "(" + calendertime + ")").getUri();
											fa = Folderuri;

										}
									} else {

										Folderuri = clientforexchange_output.createFolder(calendertime).getUri();

										fa = Folderuri;

									}

								}

							}

						}

						try {

							if (chckbx_Mail_Filter.isSelected()
									&& (mailfilterenddate == null || mailfilterstartdate == null)) {
								JOptionPane.showMessageDialog(frame, "Please Select Start and End Date",
										messageboxtitle, JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else if (chckbx_calender_box.isSelected()
									&& (Calenderfilterenddate == null || Calenderfilterstartdate == null)) {
								JOptionPane.showMessageDialog(frame, "Please Select Start and End Date",
										messageboxtitle, JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else {

								for (Map.Entry<String, List<ExchangeFolderInfo>> entry : hmexpofinal.entrySet()) {
//									if (stop) {
//									break;
//								}
									if (stop) {
										break;
									}
									long starttime = System.currentTimeMillis();
									String key = entry.getKey();
									count_destination = 0;
									try {

										try {
											clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2,
													password_p2);
											clientforexchange_input.getMailboxInfo(key);

										} catch (Exception e) {

											clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2,
													password_p2);
											userid = key;
											clientforexchange_input.impersonateUser(3, key);

										}

										f = new File(destination_path + File.separator
												+ getRidOfIllegalFileNameCharacters(key));

										destination_path = f.getAbsolutePath();
										f.mkdirs();
										pstfolderlist = new ArrayList<String>();
										listExchangdinal = new ArrayList<ExchangeFolderInfo>();
										listExchangdinal = hmexpofinal.get(key);
										pstfolderlist = hm.get(key);

										label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));

										if (filetype.equalsIgnoreCase("EML") || filetype.equalsIgnoreCase("MSG")
												|| filetype.equalsIgnoreCase("EMLX")
												|| filetype.equalsIgnoreCase("HTML")
												|| filetype.equalsIgnoreCase("MHTML")) {
											f.mkdirs();
											Convertexchange_file();

										} else if (filetype.equalsIgnoreCase("OFFICE 365")
												|| filetype.equalsIgnoreCase("HOTMAIL")
												|| filetype.equalsIgnoreCase("Live Exchange")) {
											Folderuri = clientforexchange_output.createFolder(Folderuri, key).getUri();
											Convertexchange_exchangeoutput();
											Folderuri = fa;

										} else if (filetype.equalsIgnoreCase("IMAP")
												|| filetype.equalsIgnoreCase("Hostgator email")) {
											fa = path;
											path = path + "." + key;

											clientforimap_output.createFolder(iconnforimap_output, path);
											clientforimap_output.selectFolder(iconnforimap_output, path);

											Convertexchange_imap();
											path = fa;
										} else if (filetype.equalsIgnoreCase("GMAIL")
												|| filetype.equalsIgnoreCase("AOL")
												|| filetype.equalsIgnoreCase("Icloud")
												|| filetype.equalsIgnoreCase("GoDaddy email")

												|| filetype.equalsIgnoreCase("Zoho Mail")
												|| filetype.equalsIgnoreCase("Amazon Webmail")
												|| filetype.equalsIgnoreCase("Yandex Mail")
												|| filetype.equalsIgnoreCase("YAHOO MAIL")) {
											if (filetype.equalsIgnoreCase("GoDaddy email")) {

												key = key.replaceAll("[^a-zA-Z0-9]", "");

											}
											fa = path;
											path = path + "/" + key;
											clientforimap_output.createFolder(iconnforimap_output, path);
											clientforimap_output.selectFolder(iconnforimap_output, path);
											Convertexchange_gmail();
											path = fa;
										} else if (filetype.equalsIgnoreCase("CSV")) {
											f.mkdirs();
											Convertexchange_csv();
										} else if (filetype.equalsIgnoreCase("PST")) {
											f.mkdirs();
											pst = PersonalStorage.create(
													destination_path + File.separator + key + ".pst",
													FileFormatVersion.Unicode);
											pst.getStore().changeDisplayName(key);
											Convertexchange_pstost();
											pst.dispose();
										} else if (filetype.equalsIgnoreCase("OST")) {
											f.mkdirs();
											pst = PersonalStorage.create(
													destination_path + File.separator + key + ".ost",
													FileFormatVersion.Unicode);
											pst.getStore().changeDisplayName(key);
											Convertexchange_pstost();
											pst.dispose();

										} else if (filetype.equalsIgnoreCase("MBOX")
												|| filetype.equalsIgnoreCase("Thunderbird")
												|| filetype.equalsIgnoreCase("Opera Mail")) {
											f.mkdirs();
											Convertexchange_Mbox();

										} else {
											f.mkdirs();
											Convertexchange_word();

										}
										String duration = Duration(starttime).toString();
										DefaultTableModel mode1 = (DefaultTableModel) table_fileConvertionreport_panel4
												.getModel();
										reportpath = f.getAbsolutePath();
										mode1.addRow(new Object[] { fileoption, filetype, key, Status, duration,
												count_destination, reportpath });
										destination_path = destinationfile;
										totalcount = totalcount + count_destination;
										clientforexchange_input.dispose();
									} catch (Exception e) {
										clientforexchange_input.dispose();
										continue;
									}

								}
							}

							if (filetype.equalsIgnoreCase("THUNDERBIRD")) {
								JOptionPane.showMessageDialog(frame,
										"Please open the converted file from " + destination_path + " Thunderbird",
										messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
							}
							if (filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("GMAIL")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equals("AOL")
									|| filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Zoho Mail")
									|| filetype.equalsIgnoreCase("Amazon Webmail")
									|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("hotmail")
									|| filetype.equalsIgnoreCase("IMAP")) {
								if (Desktop.isDesktopSupported()
										&& Desktop.getDesktop().isSupported(Desktop.Action.BROWSE)) {

									if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

										Desktop.getDesktop().browse(new URI("http://login.yahoo.com"));
										reportpath = "http://login.yahoo.com";
									} else if (filetype.equalsIgnoreCase("Zoho Mail")) {

										reportpath = "https://accounts.zoho.in/signin?servicename=VirtualOffice&signupurl=https://www.zoho.in/mail/zohomail-pricing.html&serviceurl=https://mail.zoho.in";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("Yandex Mail")) {

										reportpath = "https://mail.yandex.com/?uid=1213147137#tabs/relevant";
										openBrowser(reportpath);
									} else if (filetype.equalsIgnoreCase("Amazon Webmail")) {

										JOptionPane.showMessageDialog(mf,
												"Please open the converted file from  Amazon Webmail", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("GMAIL")) {

										Desktop.getDesktop().browse(new URI("https://mail.google.com"));
										reportpath = "https://mail.google.com";
									} else if (filetype.equals("AOL")) {

										Desktop.getDesktop().browse(new URI("https://login.aol.com"));
										reportpath = "https://login.aol.com";
									} else if (filetype.equalsIgnoreCase("Live Exchange")) {

										JOptionPane.showMessageDialog(frame,
												"Please open the converted file from Live Exchange", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else if (filetype.equalsIgnoreCase("Hotmail")) {

										Desktop.getDesktop().browse(new URI("https://outlook.live.com"));
										reportpath = "https://outlook.live.com";
									} else if (filetype.equalsIgnoreCase("IMAP")) {

										JOptionPane.showMessageDialog(frame,
												"Please open the converted file from  IMAP", messageboxtitle,
												JOptionPane.INFORMATION_MESSAGE);
									} else {

										Desktop.getDesktop().browse(new URI("https://outlook.office365.com"));
										reportpath = "https://outlook.office365.com";
									}
								}

							} else {

								f = new File(destination_path);

								desktop.open(f);

							}

						} catch (Error e) {
							mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());

						} finally {

							// progressBar_message_p3.setVisible(false);
							Progressbar.setVisible(false);
							Calendar cio = Calendar.getInstance();

							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_4");

							mf.logger.info("File Saved " + count_destination + System.lineSeparator() + "End Time : "
									+ cio.getTime() + System.lineSeparator()
									+ "**********************************************************");
							String duration = Duration(begintime).toString();
							DefaultTableModel mode1 = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
							mode1.addRow(new Object[] { "Total Message", "", "", "", duration, totalcount, "", "" });
						}
					}
				});

				th.start();

			}
		});
		btn_Destination.setToolTipText("Click here to Select the destination path ");
		btn_Destination.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				try {

					destinationPath();
					checkdestination = false;

				} catch (Error e) {
					mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
					e.printStackTrace();

				}

			}
		});

		panel_4 = new JPanel();
		panel_4.setBackground(Color.WHITE);
		Cardlayout.add(panel_4, "panel_4");
		panel_4.setLayout(null);

		JScrollPane scrollPane_table_panel4 = new JScrollPane();
		scrollPane_table_panel4.setBackground(Color.WHITE);
		scrollPane_table_panel4.setBounds(0, 13, 1075, 397);
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
				new String[] { "From", "To", "Mailbox id", "Status", "Duration", "Message Count", "Path" }));
		table_fileConvertionreport_panel4.getColumnModel().getColumn(0).setPreferredWidth(126);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(1).setPreferredWidth(57);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(2).setPreferredWidth(126);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(3).setPreferredWidth(35);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(4).setPreferredWidth(30);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(5).setPreferredWidth(30);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(6).setPreferredWidth(174);
		scrollPane_table_panel4.setViewportView(table_fileConvertionreport_panel4);

		btnDowloadReport = new JButton("");
		btnDowloadReport.setRolloverEnabled(false);
		btnDowloadReport.setRequestFocusEnabled(false);
		btnDowloadReport.setOpaque(false);
		btnDowloadReport.setFocusable(false);
		btnDowloadReport.setFocusTraversalKeysEnabled(false);
		btnDowloadReport.setFocusPainted(false);
		btnDowloadReport.setDefaultCapable(false);
		btnDowloadReport.setContentAreaFilled(false);
		btnDowloadReport.setBorderPainted(false);
		btnDowloadReport.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnDowloadReport.setIcon(new ImageIcon(backup.class.getResource("/download-report-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnDowloadReport.setIcon(new ImageIcon(backup.class.getResource("/download-report-btn.png")));
			}
		});

		btnDowloadReport.setIcon(new ImageIcon(backup.class.getResource("/download-report-btn.png")));
		btnDowloadReport.setToolTipText("Click here to Dowload the status report ");
		btnDowloadReport.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				cal = Calendar.getInstance();
				calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
				reportpath = logpath;
				new File(reportpath + File.separator + messageboxtitle + " report").mkdirs();

				File file = new File(reportpath + File.separator + messageboxtitle + " report" + File.separator
						+ calendertime + "report.csv");

				try {
					FileWriter outputfile = new FileWriter(file);

					CSVWriter writer = new CSVWriter(outputfile);

					String[] header = { "From", "To", "Mailbox id", "Status", "Duration", "Message Count", "Path" };

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

				} catch (Error e) {
					mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception :" + e.getMessage() + System.lineSeparator());
				}
			}
		});
		btnDowloadReport.setFont(new Font("Tahoma", Font.BOLD, 15));
		btnDowloadReport.setBounds(501, 450, 163, 39);
		panel_4.add(btnDowloadReport);

		btnConvertAgain = new JButton("");
		btnConvertAgain.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnConvertAgain.setIcon(new ImageIcon(backup.class.getResource("/convert-again-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnConvertAgain.setIcon(new ImageIcon(backup.class.getResource("/convert-again-btn.png")));
			}
		});
		btnConvertAgain.setIcon(new ImageIcon(backup.class.getResource("/convert-again-btn.png")));
		btnConvertAgain.setBorderPainted(false);
		btnConvertAgain.setContentAreaFilled(false);
		btnConvertAgain.setDefaultCapable(false);
		btnConvertAgain.setFocusTraversalKeysEnabled(false);
		btnConvertAgain.setFocusable(false);
		btnConvertAgain.setOpaque(false);
		btnConvertAgain.setRolloverEnabled(false);
		btnConvertAgain.setRequestFocusEnabled(false);
		btnConvertAgain.setToolTipText("Click here to Convert again");
		btnConvertAgain.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				filetype = "";
				path = "";
				checkconvertagain = true;
				stop = false;
				lbl_progressreport.setText("");
				btn_Destination.setEnabled(true);
				btn_previous_p3.setEnabled(true);
				btnStop.setVisible(false);
				comboBox.setVisible(false);
				label_12.setVisible(false);
				chckbx_calender_box.setEnabled(true);
				chckbxRemoveDuplicacy.setEnabled(true);
				btn_signout_p3.setVisible(false);
				dateChooser_calender_start.setEnabled(true);
				task_box.setEnabled(true);
				chckbx_Mail_Filter.setEnabled(true);
				dateChooser_calendar_end.setEnabled(true);
				dateChooser_mail_fromdate.setEnabled(true);
				dateChooser_mail_tilldate.setEnabled(true);
				dateChooser_task_start_date.setEnabled(true);
				dateChooser_task_end_date.setEnabled(true);
				comboBox_fileDestination_type.setEnabled(true);
				chckbxCustomFolderName.setEnabled(true);
				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_3");

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						comboBox_fileDestination_type.setSelectedItem("PST");
					}
				});

			}
		});
		btnConvertAgain.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnConvertAgain.setBounds(501, 568, 163, 39);
		// panel_4.add(btnConvertAgain);

		progressBar_message_p3 = new JProgressBar();
		progressBar_message_p3.setBounds(0, 0, 11, 0);
		panel_4.add(progressBar_message_p3);
		progressBar_message_p3.setBackground(Color.WHITE);

		label_10 = new JLabel("");
		label_10.setIcon(new ImageIcon(backup.class.getResource("/bottom.png")));
		label_10.setBounds(0, 556, 1075, 64);
		panel_4.add(label_10);

		JLabel lblNewLabel_5 = new JLabel("");
		lblNewLabel_5.setIcon(new ImageIcon(backup.class.getResource("/topbar.png")));
		lblNewLabel_5.setBounds(0, 0, 1075, 71);
		contentPane.add(lblNewLabel_5);

	}

	void destinationPath() throws Exception {
		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

		jFileChooser.setMultiSelectionEnabled(true);

		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		jFileChooser.showOpenDialog(backup.this);

		File file = jFileChooser.getSelectedFile();

		String destination = file.getAbsolutePath();

		tf_Destination_Location.setText(destination);

	}

	public void readexchangePublicB_R(DefaultMutableTreeNode node, String userid) {
		String rootUri = clientforexchange_input.getMailboxInfo().getRootUri();

		listExchangemesingos = new ArrayList<ExchangeFolderInfo>();
		listFolderinfostring = new ArrayList<String>();
		ExchangeFolderInfoCollection folderInfoCollection = null;
		try {
			folderInfoCollection = clientforexchange_input.listPublicFolders();
		} catch (Exception e1) {
			e1.printStackTrace();

		}

		int co = folderInfoCollection.size();

		if (co == 0) {
			JOptionPane.showMessageDialog(backup.this, "  There are no folder in public folder of " + userid + "!!",
					messageboxtitle, JOptionPane.ERROR_MESSAGE,
					new ImageIcon(backup.class.getResource("/information.png")));

			CardLayout card = (CardLayout) Cardlayout.getLayout();

			card.show(Cardlayout, "panel_tablecontentBR");

		} else {
			for (int j = 0; j < co; j++) {
				try {

					if (stop_tree) {
						break;
					}

					ExchangeFolderInfo folderInfo = folderInfoCollection.get_Item(j);
					countf++;
					if (!(folderInfo.getDisplayName().equalsIgnoreCase("PersonMetadata"))) {
						DefaultMutableTreeNode t = new DefaultMutableTreeNode(
								"<html><b>" + folderInfo.getDisplayName());
						countf++;
						node.add(t);
						String s = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
						listExchangemesingos.add(folderInfo);
						listFolderinfostring.add(s);

						// obTh.ob.MessageLabel.setText(folderInfo.getDisplayName());
						// System.out.println(s);

						if (folderInfo.getChildFolderCount() > 0) {

							readexchange_subfolder(folderInfo, t, s, rootUri, userid);

						}
					}
				} catch (Exception e) {
					e.printStackTrace();
					continue;
				}

			}
		}
		System.out.println(userid);
		hmstring.put(userid, listFolderinfostring);
		hmexpo.put(userid, listExchangemesingos);
	}

	public void readexchangArchieve(DefaultMutableTreeNode node, String userid) {

		ExchangeFolderInfo folderj = null;
		try {
			folderj = clientforexchange_input.getFolderInfo("ArchiveMsgFolderRoot");
		} catch (Exception e2) {

			JOptionPane.showMessageDialog(backup.this,
					"  The Archieve folder is disable for this user " + userid + "!!", messageboxtitle,
					JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));

			CardLayout card = (CardLayout) Cardlayout.getLayout();

			card.show(Cardlayout, "panel_tablecontentBR");

			openBrowser(
					"https://docs.microsoft.com/en-us/microsoft-365/compliance/enable-archive-mailboxes?view=o365-worldwide");
		}
		String rootUri = folderj.getUri();

		listExchangemesingos = new ArrayList<ExchangeFolderInfo>();
		listFolderinfostring = new ArrayList<String>();
		ExchangeFolderInfoCollection folderInfoCollection = null;
		try {
			folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);
		} catch (Exception e1) {
			CardLayout card = (CardLayout) Cardlayout.getLayout();

			card.show(Cardlayout, "panel_tablecontentBR");

			openBrowser(
					"https://docs.microsoft.com/en-us/microsoft-365/compliance/enable-archive-mailboxes?view=o365-worldwide");
		}
		// System.out.println(folderInfoCollection.size());

		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {

				if (stop_tree) {
					break;
				}

				ExchangeFolderInfo folderInfo = folderInfoCollection.get_Item(j);

				System.out.println(folderInfo.getDisplayName());

				if (!(folderInfo.getDisplayName().equalsIgnoreCase("PersonMetadata"))) {
					DefaultMutableTreeNode t = new DefaultMutableTreeNode("<html><b>" + folderInfo.getDisplayName());

					node.add(t);
					countf++;
					countf++;
					String s = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
					listExchangemesingos.add(folderInfo);
					listFolderinfostring.add(s);

					// obTh.ob.MessageLabel.setText(folderInfo.getDisplayName());
					// System.out.println(s);

					if (folderInfo.getChildFolderCount() > 0) {

						readexchange_subfolder(folderInfo, t, s, rootUri, userid);

					}
				}
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}

		}
		System.out.println(userid);
		hmstring.put(userid, listFolderinfostring);
		hmexpo.put(userid, listExchangemesingos);
	}

	public void readexchangeB_R(DefaultMutableTreeNode node, String userid) {

		String rootUri = null;
		try {

			ExchangeMailboxInfo m = clientforexchange_input.getMailboxInfo(userid);
			rootUri = m.getRootUri();

		} catch (Exception e1) {

			e1.printStackTrace();
			clientforexchange_input = EWSClient.getEWSClient("https://outlook.office365.com/EWS/Exchange.asmx",
					username_p2, password_p2);
			clientforexchange_input.impersonateUser(3, userid);
			ExchangeMailboxInfo m = clientforexchange_input.getMailboxInfo();
			rootUri = m.getRootUri();
		}

		listExchangemesingos = new ArrayList<ExchangeFolderInfo>();
		listFolderinfostring = new ArrayList<String>();

		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);

		int co = folderInfoCollection.size();

		if (co == 0) {
			JOptionPane.showMessageDialog(backup.this, "  Cannot Acces  " + userid + "!!", messageboxtitle,
					JOptionPane.ERROR_MESSAGE, new ImageIcon(backup.class.getResource("/information.png")));

			CardLayout card = (CardLayout) Cardlayout.getLayout();

			card.show(Cardlayout, "panel_tablecontentBR");

		} else {

			for (int j = 0; j < folderInfoCollection.size(); j++) {
				try {

					if (stop_tree) {
						break;
					}

					ExchangeFolderInfo folderInfo = folderInfoCollection.get_Item(j);

					if (!(folderInfo.getDisplayName().equalsIgnoreCase("PersonMetadata"))) {
						DefaultMutableTreeNode t = new DefaultMutableTreeNode(
								"<html><b>" + folderInfo.getDisplayName());
						countf++;
						node.add(t);
						String s = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
						listExchangemesingos.add(folderInfo);
						listFolderinfostring.add(s);

						// obTh.ob.MessageLabel.setText(folderInfo.getDisplayName());
						// System.out.println(s);

						if (folderInfo.getChildFolderCount() > 0) {

							readexchange_subfolder(folderInfo, t, s, rootUri, userid);

						}
					}
				} catch (Exception e) {
					e.printStackTrace();
					continue;
				}

			}
		}
		System.out.println(userid);
		hmstring.put(userid, listFolderinfostring);
		hmexpo.put(userid, listExchangemesingos);
	}

	public void readexchange_subfolder(ExchangeFolderInfo folderInfo, DefaultMutableTreeNode node, String s,
			String rootUri, String userid) {
		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri,
				folderInfo.getUri());
		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {

				if (stop_tree) {
					break;
				}
				ExchangeFolderInfo folderInfo1 = folderInfoCollection.get_Item(j);

				DefaultMutableTreeNode t = new DefaultMutableTreeNode("<html><b>" + folderInfo1.getDisplayName());

				node.add(t);
				// obTh.ob.MessageLabel.setText(folderInfo1.getDisplayName());

				s = s + File.separator + folderInfo1.getDisplayName().replaceAll("[\\[\\]]", "");
				listExchangemesingos.add(folderInfo1);
				listFolderinfostring.add(s);

				if (folderInfo1.getChildFolderCount() > 0) {

					readexchange_subfolder(folderInfo1, t, s, rootUri, userid);

				}
				s = s.replace(File.separator + folderInfo1.getDisplayName().replaceAll("[\\[\\]]", ""), "");
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}

		}

	}

	public void Convertexchange_file() {

		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				String folderuri = folderInfo.getUri();
				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				path = pstfolderlist.get(l);

				new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();

				int totalMessages = folderInfo.getTotalCount();
				System.out.println(totalMessages);
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages();
						}

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri(), folderuri);

									Date reciveddate = appointment.getStartDate();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											appointment.save(
													destination_path + File.separator + path + File.separator + i
															+ getRidOfIllegalFileNameCharacters(
																	calendarname(appointment))
															+ ".ics",
													AppointmentSaveFormat.Ics);

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											appointment.save(
													destination_path + File.separator + path + File.separator + i
															+ getRidOfIllegalFileNameCharacters(
																	calendarname(appointment))
															+ ".ics",
													AppointmentSaveFormat.Ics);

										}
										count_destination++;
									} else {

										appointment.save(destination_path + File.separator + path + File.separator + i
												+ getRidOfIllegalFileNameCharacters(calendarname(appointment)) + ".ics",
												AppointmentSaveFormat.Ics);
										count_destination++;

									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());

									MailMessage msg2 = taskMessage.toMailMessage(new MailConversionOptions());
									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage eml = msg.toMailMessage(options);
									Date reciveddate = eml.getDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											Mailmessage_file(eml);

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											Mailmessage_file(eml);

										}
									} else {
										Mailmessage_file(eml);
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());
									contact.save(
											destination_path + File.separator + path + File.separator + i
													+ getRidOfIllegalFileNameCharacters(contactname(contact)) + ".vcf",
											ContactSaveFormat.VCard);

									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailfile(msg1, reciveddate);

								} else {
									MailMessage msg2 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());

									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage msg1 = msg.toMailMessage(options);
									Date reciveddate = msg1.getDate();
									mailfile mf1 = new mailfile(mf, filetype, destination_path, path, msg1, reciveddate,
											chckbx_Mail_Filter.isSelected(), mailfilterstartdate, mailfilterenddate,
											filepath);
									Thread saveTh = new Thread(mf1);
									saveTh.start();
									saveTh.join();

								}

								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}

				}
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
	}

	public void Convertexchange_csv() {
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				path = pstfolderlist.get(l);

				new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();

				String pa = getRidOfIllegalFileNameCharacters(folderInfo.getDisplayName());
				File file1 = new File(destination_path + File.separator + path + File.separator + pa + ".csv");

				FileWriter outputfile = new FileWriter(file1);

				CSVWriter writer = new CSVWriter(outputfile);

				String[] header = { "Date", "Subject", "Body", "From", "To", "CC:" };

				writer.writeNext(header);

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages();
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());

									Date reciveddate = appointment.getStartDate();

									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + i + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + i + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + calendarname(appointment) + ".ics"));
									file.delete();
									MapiMessage msg = MapiMessage.fromMailMessage(mapi,
											MapiConversionOptions.getASCIIFormat());
									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											mapicsv(msg, reciveddate, writer);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {
											mapicsv(msg, reciveddate, writer);

											count_destination++;

										}

									} else {

										mapicsv(msg, reciveddate, writer);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + i + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);
									Date reciveddate = null;
									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + i + contactname(contact) + ".vcf");
									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + contactname(contact) + ".vcf"));
									file.delete();
									MapiMessage msg = MapiMessage.fromMailMessage(mapi,
											MapiConversionOptions.getASCIIFormat());

									mapicsv(msg, reciveddate, writer);
									// System.out.println("Done");
									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									MailMessage eml = taskMessage.toMailMessage(new MailConversionOptions());
									Date reciveddate = eml.getDate();
									MapiMessage msg = MapiMessage.fromMailMessage(eml,
											MapiConversionOptions.getASCIIFormat());
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											mapicsv(msg, reciveddate, writer);
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											mapicsv(msg, reciveddate, writer);
											count_destination++;

										}
									} else {
										mapicsv(msg, reciveddate, writer);
										count_destination++;
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									MapiMessage msg = MapiMessage.fromMailMessage(msg1,
											MapiConversionOptions.getASCIIFormat());
									Date reciveddate = msg1.getDate();
									mapicsv(msg, reciveddate, writer);

								} else {
									MailMessage msg2 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage msg1 = msg.toMailMessage(options);
									Date reciveddate = msg1.getDate();
									mapicsv(msg, reciveddate, writer);
								}

								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {

								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						// System.out.println(pageInfo.getLastPage());
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}
					writer.close();
				}
			} catch (Exception e) {
				continue;
			}
		}

	}

	public void Convertexchange_word() {
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				path = pstfolderlist.get(l);

				new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}

						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages();
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());

									Date reciveddate = appointment.getStartDate();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											appointment.save(
													destination_path + File.separator + path + File.separator + i
															+ getRidOfIllegalFileNameCharacters(
																	calendarname(appointment))
															+ ".ics",
													AppointmentSaveFormat.Ics);

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											appointment.save(
													destination_path + File.separator + path + File.separator + i
															+ getRidOfIllegalFileNameCharacters(
																	calendarname(appointment))
															+ ".ics",
													AppointmentSaveFormat.Ics);

										}
										count_destination++;
									} else {

										appointment.save(destination_path + File.separator + path + File.separator + i
												+ getRidOfIllegalFileNameCharacters(calendarname(appointment)) + ".ics",
												AppointmentSaveFormat.Ics);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());

									MailMessage msg2 = taskMessage.toMailMessage(new MailConversionOptions());
									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage eml = msg.toMailMessage(options);
									Date reciveddate = eml.getDate();
									MailWord mf1 = new MailWord(mf, filetype, destination_path, path, temppath, eml,
											reciveddate, task_box.isSelected(), taskfilterstartdate, taskfilterenddate);

									Thread saveTh = new Thread(mf1);
									saveTh.start();
									saveTh.join();

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());
									contact.save(
											destination_path + File.separator + path + File.separator + i
													+ getRidOfIllegalFileNameCharacters(contactname(contact)) + ".vcf",
											ContactSaveFormat.VCard);
									// System.out.println("Done");
									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									MailWord mf1 = new MailWord(mf, filetype, destination_path, path, temppath, msg1,
											reciveddate, task_box.isSelected(), taskfilterstartdate, taskfilterenddate);

									Thread saveTh = new Thread(mf1);
									saveTh.start();
									saveTh.join();

									if (chckbx_Mail_Filter.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {
											count_destination++;

										}
									} else {
										count_destination++;

									}

								} else {
									MailMessage msg2 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage msg1 = msg.toMailMessage(options);
									Date reciveddate = msg1.getDate();
									MailWord mf1 = new MailWord(mf, filetype, destination_path, path, temppath, msg1,
											reciveddate, chckbx_Mail_Filter.isSelected(), mailfilterstartdate,
											mailfilterenddate);

									Thread saveTh = new Thread(mf1);
									saveTh.start();
									saveTh.join();
									
									if (chckbx_Mail_Filter.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {
											count_destination++;

										}
									} else {
										count_destination++;

									}

								}

								// int messageprogress = (count * 100) / messagesize;
								// iprogressBar_message_p3.setValue(messageprogress);
								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								} else {
									// System.out.println(">>>>>>>>>>>> " + e.getMessage());
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						// System.out.println(pageInfo.getLastPage());
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}

				}
			} catch (Exception e) {
				continue;
			}
		}
	}

	public void Convertexchange_Mbox() {
		String[] fromMarker = { null };
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}

				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				path = pstfolderlist.get(l);

				new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();
				MboxrdStorageWriter wr = null;
				String pa = getRidOfIllegalFileNameCharacters(folderInfo.getDisplayName());
				if (filetype.equalsIgnoreCase("Opera Mail")) {
					wr = new MboxrdStorageWriter(
							destination_path + File.separator + path + File.separator + pa + ".mbs", false);
				} else if (filetype.equalsIgnoreCase("Thunderbird")) {
					wr = new MboxrdStorageWriter(destination_path + File.separator + path + File.separator + pa, false);
				} else {
					wr = new MboxrdStorageWriter(
							destination_path + File.separator + path + File.separator + pa + ".mbx", false);
				}

				// writer.writeNext(header);

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages();
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}
								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());

									Date reciveddate = appointment.getStartDate();

									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + i + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + i + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + calendarname(appointment) + ".ics"));
									file.delete();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											wr.writeMessage(mapi, fromMarker);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											wr.writeMessage(mapi, fromMarker);
											count_destination++;

										}

									} else {

										wr.writeMessage(mapi, fromMarker);
										count_destination++;

									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + i + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);

									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + i + contactname(contact) + ".vcf");
									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + contactname(contact) + ".vcf"));
									file.delete();

									wr.writeMessage(mapi, fromMarker);

									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									MailMessage eml = taskMessage.toMailMessage(new MailConversionOptions());
									Date reciveddate = eml.getDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											wr.writeMessage(eml, fromMarker);
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											wr.writeMessage(eml, fromMarker);
											count_destination++;

										}
									} else {
										wr.writeMessage(eml, fromMarker);
										count_destination++;
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailmbox(msg1, reciveddate, wr);

								} else {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailmbox(msg1, reciveddate, wr);

								}

								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						// System.out.println(pageInfo.getLastPage());
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}
					wr.dispose();
				}
			} catch (Exception e) {
				continue;
			}
		}
	}

	public void Convertexchange_pstost() {
		pst.createPredefinedFolder("Contact", StandardIpmFolder.Contacts, true);
		pst.createPredefinedFolder("Calender", StandardIpmFolder.Appointments, true);
		pst.createPredefinedFolder("Task", StandardIpmFolder.Tasks, true);
		List<String> ss = new ArrayList<String>();
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}

				if (ss.contains(pstfolderlist.get(l))) {

					continue;

				} else {
					ss.add(pstfolderlist.get(l));
				}

				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				path = pstfolderlist.get(l);

				// new File(destination_path + File.separator + path).mkdirs();

				// listduplicacy.clear();

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
//						clientforexchange_input.se
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							System.out.println();
							msgInfoColl = clientforexchange_input.listMessages();

							System.out.println(msgInfoColl.size());
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());
									MailMessage msg = new MailMessage();
									msg.addAlternateView(appointment.requestApointment());
									MapiMessage mapi = MapiMessage.fromMailMessage(msg);
									MapiCalendar calendar = (MapiCalendar) mapi.toMapiMessageItem();

									info1 = pst.getPredefinedFolder(StandardIpmFolder.Appointments);
									Date reciveddate = appointment.getStartDate();
									// MapiMessage MapiMessage =
									// clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											info1.addMapiMessageItem(calendar);
											count_destination++;
										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											info1.addMapiMessageItem(calendar);
											count_destination++;
										}

									} else {

										info1.addMapiMessageItem(calendar);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									if (pst.getRootFolder().getSubFolder(path) == null) {
										info = pst.getRootFolder().addSubFolder(path, true);
									} else {
										info = pst.getRootFolder().getSubFolder(path, true);
									}
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									info1 = pst.getPredefinedFolder(StandardIpmFolder.Tasks);
									MailMessage msg2 = taskMessage.toMailMessage(new MailConversionOptions());
									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage eml = msg.toMailMessage(options);
									Date reciveddate = eml.getDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											info1.addMessage(msg);

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {
											info1.addMessage(msg);

										}
									} else {
										info1.addMessage(msg);
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());
									contact.save(temppath + File.separator + 1 + "contacttemp.vcf",
											ContactSaveFormat.VCard);
									MapiContact con = Contact.to_MapiContact(contact);

									info1 = pst.getPredefinedFolder(StandardIpmFolder.Contacts);
									info1.addMapiMessageItem(con);
									f.delete();
									count_destination++;

								} else {
									if (pst.getRootFolder().getSubFolder(path) == null) {
										info = pst.getRootFolder().addSubFolder(path, true);
									} else {
										info = pst.getRootFolder().getSubFolder(path, true);
									}
									MailMessage msg2 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());

									MapiMessage msg = MapiMessage.fromMailMessage(msg2,
											MapiConversionOptions.getASCIIFormat());
									MailConversionOptions options = new MailConversionOptions();
									MailMessage msg1 = msg.toMailMessage(options);
									Date reciveddate = msg1.getDate();

									if (chckbx_Mail_Filter.isSelected()) {
										if (reciveddate.after(mailfilterstartdate)
												&& reciveddate.before(mailfilterenddate)) {

											if (msgInfo.isRead()) {
												info.addMessage(msg);

											} else {
												info.addMessage(msg);

												msg.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);

												clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

											}

											count_destination++;

										} else if (reciveddate.equals(mailfilterstartdate)
												|| reciveddate.equals(mailfilterenddate)) {
											if (msgInfo.isRead()) {
												info.addMessage(msg);

											} else {
												info.addMessage(msg);

												msg.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);

												clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

											}
											count_destination++;
										}
									} else {
										if (msgInfo.isRead()) {
											info.addMessage(msg);

										} else {
											info.addMessage(msg);

											msg.setMessageFlags(MapiMessageFlags.MSGFLAG_NOTIFYUNREAD);

											clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

										}
										count_destination++;
									}

								}

								// int messageprogress = (count * 100) / messagesize;
								// iprogressBar_message_p3.setValue(messageprogress);
								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {

								e.printStackTrace();
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								} else {
									// System.out.println(">>>>>>>>>>>> " + e.getMessage());
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
//						//						System.out.println(pageInfo.getLastPage());
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}

				}
			} catch (Exception e) {
				continue;
			}
		}
		pst.dispose();
	}

	public void Convertexchange_imap() {
		String pa = "";
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

				pa = pstfolderlist.get(l);
				pa = pa.replace(File.separator, ".");

				// new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();

				path = path + "." + pa;
				clientforimap_output.createFolder(iconnforimap_output, path);
				clientforimap_output.selectFolder(iconnforimap_output, path);

				// writer.writeNext(header);

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}
//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages();
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);

//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());

									Date reciveddate = appointment.getStartDate();

									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + i + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + i + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + calendarname(appointment) + ".ics"));
									file.delete();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
											count_destination++;

										}

									} else {

										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + i + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);

									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + i + contactname(contact) + ".vcf");
									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + contactname(contact) + ".vcf"));
									file.delete();

									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);

									// System.out.println("Done");
									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									MailMessage eml = taskMessage.toMailMessage(new MailConversionOptions());
									Date reciveddate = eml.getDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
											count_destination++;

										}
									} else {
										clientforimap_output.appendMessage(iconnforimap_output, path, eml);
										count_destination++;
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailimap(msg1, reciveddate, path);

								} else {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailimap(msg1, reciveddate, path);

								}
								// int messageprogress = (count * 100) / messagesize;
								// iprogressBar_message_p3.setValue(messageprogress);
								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								} else {
									// System.out.println(">>>>>>>>>>>> " + e.getMessage());
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						// System.out.println(pageInfo.getLastPage());
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {
						e.printStackTrace();
						continue;
					}

				}
			} catch (Exception e) {
				continue;
			}
		}
	}

	public void Convertexchange_gmail() {
		String pa = "";
		for (int l = 0; l < pstfolderlist.size(); l++) {
			try {
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}

				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());

				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());
				pa = pstfolderlist.get(l);
				pa = pa.replace(File.separator, "/");

				// new File(destination_path + File.separator + path).mkdirs();

				listduplicacy.clear();
				if (filetype.equalsIgnoreCase("GoDaddy email")) {
					pa = pa.replaceAll("[^a-zA-Z0-9/]", "");

				}

				path = path + "/" + pa;
				clientforimap_output.createFolder(iconnforimap_output, path);
				clientforimap_output.selectFolder(iconnforimap_output, path);

				// writer.writeNext(header);

				int totalMessages = folderInfo.getTotalCount();
				int messagesPerPage = 500;
				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
				listduplicacy.clear();
				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);

				for (int i1 = 1; i1 <= totalPages; i1++) {
					try {
						if (demo) {
							if (i1 > 1) {
								break;
							}
						}

//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						ExchangeMessagePageInfo pageInfo = null;
						ExchangeMessageInfoCollection msgInfoColl = null;
						try {
							pageInfo = clientforexchange_input.listMessagesByPage(folderInfo.getUri(), messagesPerPage,
									i1 - 1);
							msgInfoColl = pageInfo.getItems();
						} catch (Exception e1) {
							msgInfoColl = clientforexchange_input.listMessages(folderInfo.getUri());

							System.out.println(msgInfoColl.size());
						}

						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());

						int messagesize;

						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {
								if ((i % 100) == 0) {
									System.gc();
								}

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);
//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								// int count = 1;

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());

									Date reciveddate = appointment.getStartDate();

									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + i + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + i + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + calendarname(appointment) + ".ics"));
									file.delete();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
											count_destination++;

										}

									} else {

										clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + i + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);

									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + i + contactname(contact) + ".vcf");
									mapi.addAttachment(new Attachment(
											temppath + File.separator + i + contactname(contact) + ".vcf"));
									file.delete();

									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);

									// System.out.println("Done");
									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									MailMessage eml = taskMessage.toMailMessage(new MailConversionOptions());
									Date reciveddate = eml.getDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
											count_destination++;

										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
											count_destination++;

										}
									} else {
										clientforimap_output.appendMessage(iconnforimap_output, path, eml);
										count_destination++;
									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailimap(msg1, reciveddate, path);

								} else {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									mailimap(msg1, reciveddate, path);

								}

								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());

							} catch (Exception e) {

								if (e.getMessage().equalsIgnoreCase(
										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
										|| e.getMessage().contains("ConnectFailure")
										|| e.getMessage().contains("Rate limit hit")
										|| e.getMessage().contains("Operation failed")
										|| e.getMessage().contains(
												"No connection could be made because the target machine actively refused it.")
										|| e.getMessage().contains("Operation has been canceled")
										|| e.getMessage().contains(
												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
									Progressbar.setVisible(false);
									connectionHandle(e.getMessage());
									i--;
								}
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());

								continue;
							}

						}
						if (pageInfo != null) {
							if (pageInfo.getLastPage()) {

								break;
							}
						}

					} catch (Exception e) {

						continue;
					}

				}
			} catch (Exception e) {
				continue;
			}
		}
	}

//	public void Convertexchange_exchangeoutput() {
//		String pa = "";
//		String rootUri = clientforexchange_input.getMailboxInfo().getRootUri();
//		Calendar cal = Calendar.getInstance();
//		String time = cal.getTime().toString();
//		String parentfolder = clientforexchange_output.createFolder(time).getUri();
//		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);
//		String folderuri = "";
//		List<String> listdupliccontact = new ArrayList<String>();
//		if (chckbxCustomFolderName.isSelected()) {
//			String customerfolder = textField_customfolder.getText().replace("//s", "");
//
//			customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
//			try {
//
//				folderuri = clientforexchange_output.createFolder(customerfolder).getUri();
//
//				path = customerfolder;
//
//			} catch (Exception e) {
//
//				folderuri = clientforexchange_output.createFolder(customerfolder + "(" + calendertime + ")").getUri();
//				path = customerfolder + "(" + calendertime + ")";
//			}
//		} else {
//
//			folderuri = clientforexchange_output.createFolder(calendertime).getUri();
//
//			folderuri = clientforexchange_output.createFolder(folderuri, textField_username_p2.getText()).getUri();
//			path = textField_username_p2.getText();
//		}
//		HashMap<String, String> exchangepath = new HashMap<String, String>();
//		exchangepath.put(path, folderuri);
//		final String mainfolderpath = path;
//		for (int l = 0; l < pstfolderlist.size(); l++) {
//			try {
////				if (stop) {
////				break;
////			}
//				if (stop) {
//					break;
//				}
//
//				ExchangeFolderInfo folderInfo = clientforexchange_input.getFolderInfo(listExchangdinal.get(l).getUri());
//
//				lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());
//				pa = pstfolderlist.get(l);
//				pa = pa.replace(File.separator, "/");
//
//				// new File(destination_path + File.separator + path).mkdirs();
//
//				listduplicacy.clear();
//				if (filetype.equalsIgnoreCase("GoDaddy email")) {
//					pa = pa.replaceAll("[^a-zA-Z0-9/]", "");
//
//				}
//
//				path = path + "/" + pa;
//				folderuri = exchangepath.get(mainfolderpath);
//				String[] suv =pstfolderlist.get(l).split("/");
//
//				String s="";
//					for (int i = 0; i < suv.length; i++) {
//						if (i == 0) {
//							s = getRidOfIllegalFileNameCharacters(suv[i]);
//							pa = s;
//
//							if (!listdupliccontact.contains(pa)) {
//								listdupliccontact.add(pa);
//								try {
//									folderuri = clientforexchange_output.createFolder(folderuri, s).getUri();
//									exchangepath.put(pa, folderuri);
//								} catch (Exception e) {
//
//								}
//							} else {
//								folderuri = exchangepath.get(pa);
//							}
//						} else {
//							s = getRidOfIllegalFileNameCharacters(suv[i]);
//							pa = pa + "/" + s;
//							if (!listdupliccontact.contains(pa)) {
//								listdupliccontact.add(pa);
//								try {
//									folderuri = clientforexchange_output.createFolder(folderuri, s).getUri();
//									exchangepath.put(pa, folderuri);
//								} catch (Exception e) {
//								}
//							} else {
//								folderuri = exchangepath.get(pa);
//							}
//						}
//
//					}
//				
//
//				// writer.writeNext(header);
//
//				int totalMessages = folderInfo.getTotalCount();
//				int messagesPerPage = 500;
//				int totalPages = (int) Math.ceil(((double) totalMessages / (double) messagesPerPage));
//				listduplicacy.clear();
//				System.out.println(folderInfo.getDisplayName() + " size of main folder : " + totalMessages);
//
//				for (int i1 = 1; i1 <= totalPages; i1++) {
//					try {
//						if (demo) {
//							if (i1 > 1) {
//								break;
//							}
//						}
//
////						if (stop) {
////						break;
////					}
//						if (stop) {
//							break;
//						}
//						ExchangeMessagePageInfo pageInfo = clientforexchange_input
//								.listMessagesByPage(folderInfo.getUri(), messagesPerPage, i1 - 1);
//
//						ExchangeMessageInfoCollection msgInfoColl = pageInfo.getItems();
//
//						System.out.println("Page No " + i1 + " Number of folder : " + msgInfoColl.size());
//
//						int messagesize;
//
//						if (demo) {
//							if (msgInfoColl.size() <= All_Data.demo_count) {
//								messagesize = msgInfoColl.size();
//							} else {
//								messagesize = All_Data.demo_count;
//							}
//
//						} else {
//							messagesize = msgInfoColl.size();
//						}
//
//						for (int i = 0; i < messagesize; i++) {
//
//							try {
//								if ((i % 100) == 0) {
//									System.gc();
//								}
//
//								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);
////								if (stop) {
////								break;
////							}
//								if (stop) {
//									break;
//								}
//
//								// int count = 1;
//
//								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
//									Appointment appointment = clientforexchange_input
//											.fetchAppointment(msgInfo.getUniqueUri());
//									if(clientforexchange_output.folderExists(pa, s, null)) {
//										
//										
//									}
//
//									Date reciveddate = appointment.getStartDate();
//
//									MailMessage mapi = new MailMessage();
//
//									appointment.save(temppath + File.separator + i + calendarname(appointment) + ".ics",
//											AppointmentSaveFormat.Ics);
//									File file = new File(
//											temppath + File.separator + i + calendarname(appointment) + ".ics");
//
//									mapi.addAttachment(new Attachment(
//											temppath + File.separator + i + calendarname(appointment) + ".ics"));
//									file.delete();
//
//									if (chckbx_calender_box.isSelected()) {
//										if (reciveddate.after(Calenderfilterstartdate)
//												&& reciveddate.before(Calenderfilterenddate)) {
//											clientforexchange_output.appendMessage(folderuri, mapi);
//											count_destination++;
//
//										}
//
//										else if (reciveddate.equals(Calenderfilterstartdate)
//												|| reciveddate.equals(Calenderfilterenddate)) {
//
//											clientforexchange_output.appendMessage(folderuri, mapi);
//											count_destination++;
//										}
//
//									} else {
//										clientforexchange_output.appendMessage(folderuri, mapi);
//										count_destination++;
//
//									}
//
//									// System.out.println("Done");
//								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
//									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());
//
//									contact.save(temppath + File.separator + i + contactname(contact) + ".vcf",
//											ContactSaveFormat.VCard);
//
//									MailMessage mapi = new MailMessage();
//									File file = new File(temppath + File.separator + i + contactname(contact) + ".vcf");
//									mapi.addAttachment(new Attachment(
//											temppath + File.separator + i + contactname(contact) + ".vcf"));
//									file.delete();
//
//									clientforimap_output.appendMessage(iconnforimap_output, path, mapi);
//
//						
//									count_destination++;
//
//								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
//									MapiMessage taskMessage = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
//									MailMessage eml = taskMessage.toMailMessage(new MailConversionOptions());
//									Date reciveddate = eml.getDate();
//									if (task_box.isSelected()) {
//										if (reciveddate.after(taskfilterstartdate)
//												&& reciveddate.before(taskfilterenddate)) {
//											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
//											count_destination++;
//
//										} else if (reciveddate.equals(taskfilterstartdate)
//												|| reciveddate.equals(taskfilterenddate)) {
//
//											clientforimap_output.appendMessage(iconnforimap_output, path, eml);
//											count_destination++;
//
//										}
//									} else {
//										clientforimap_output.appendMessage(iconnforimap_output, path, eml);
//										count_destination++;
//									}
//
//								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
//										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
//									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
//									Date reciveddate = msg1.getDate();
//									mailimap(msg1, reciveddate, path);
//
//								} else {
//									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
//									Date reciveddate = msg1.getDate();
//									mailimap(msg1, reciveddate, path);
//
//								}
//
//								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
//										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());
//
//							} catch (Exception e) {
//
//								if (e.getMessage().equalsIgnoreCase(
//										"The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
//										|| e.getMessage().contains("ConnectFailure")
//										|| e.getMessage().contains("Rate limit hit")
//										|| e.getMessage().contains("Operation failed")
//										|| e.getMessage().contains(
//												"No connection could be made because the target machine actively refused it.")
//										|| e.getMessage().contains("Operation has been canceled")
//										|| e.getMessage().contains(
//												"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
//									Progressbar.setVisible(false);
//									connectionHandle(e.getMessage());
//									i--;
//								}
//								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());
//
//								continue;
//							}
//
//						}
//						if (pageInfo.getLastPage()) {
//
//							break;
//						}
//
//					} catch (Exception e) {
//
//						continue;
//					}
//
//				}
//			} catch (Exception e) {
//				continue;
//			}
//		}
//	}
//	

	public void Convertexchange_exchangeoutput() {
		// ExchangeMailboxInfo mailboxInfo = clientforexchange_input.getMailboxInfo();

		String rootUri = clientforexchange_input.getMailboxInfo().getRootUri();
		Calendar cal = Calendar.getInstance();
		String time = cal.getTime().toString();
		String parentfolder = clientforexchange_output.createFolder(time).getUri();
		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);

		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {
				ExchangeFolderInfo folderInfo = folderInfoCollection.get_Item(j);
//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				String subfolder = parentfolder;

				System.out.println(folderInfo.getDisplayName());

				if (pstfolderlist.size() == countcheck) {
					break;
				}
				lbl_progressreport.setText("Getting Folder please wait ");
				for (int l = 0; l < pstfolderlist.size(); l++) {
//					if (stop) {
//					break;
//				}
					if (stop) {
						break;
					}
					if (listExchangdinal.get(l).getUri().equalsIgnoreCase(folderInfo.getUri())) {
						countcheck++;
						lbl_progressreport.setText("Getting Folder : " + folderInfo.getDisplayName());

						lbl_progressreport.setText("Getting : " + Folder);
//						if (stop) {
//						break;
//					}
						if (stop) {
							break;
						}
						Folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
						listduplicacy.clear();
						lbl_progressreport.setText("Getting : " + Folder);

						subfolder = clientforexchange_output
								.createFolder(parentfolder, Folder, folderInfo.getFolderType()).getUri();

						// progressBar_message_p3.setStringPainted(true);
						// progressBar_message_p3.setValue(0);
						ExchangeMessageInfoCollection msgInfoColl = clientforexchange_input.listMessages();
						int messagesize;
						if (demo) {
							if (msgInfoColl.size() <= All_Data.demo_count) {
								messagesize = msgInfoColl.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = msgInfoColl.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {

								ExchangeMessageInfo msgInfo = msgInfoColl.get_Item(i);
//								if (stop) {
//								break;
//							}
								if (stop) {
									break;
								}

								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());
									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + calendarname(appointment) + ".ics"));
									file.delete();

									Date reciveddate = appointment.getStartDate();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {
											clientforexchange_output.appendMessage(subfolder, mapi);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											clientforexchange_output.appendMessage(subfolder, mapi);
											count_destination++;
										}

									} else {
										clientforexchange_output.appendMessage(subfolder, mapi);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);

									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + contactname(contact) + ".vcf");
									mapi.addAttachment(
											new Attachment(temppath + File.separator + contactname(contact) + ".vcf"));
									file.delete();

									clientforexchange_output.appendMessage(subfolder, mapi);

									count_destination++;

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.MailMessage
										|| msgInfo.getMessageInfoType() == MessageInfoType.Post) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();

									String messageid = mailexchange(msg1, reciveddate, clientforexchange_output,
											subfolder);

									if (msgInfo.isRead()) {

										clientforexchange_output.setReadFlag(messageid, true);

									} else {
										clientforexchange_output.setReadFlag(messageid, false);
										clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.None) {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									String messageid = mailexchange(msg1, reciveddate, clientforexchange_output,
											subfolder);

									if (msgInfo.isRead()) {

										clientforexchange_output.setReadFlag(messageid, true);

									} else {
										clientforexchange_output.setReadFlag(messageid, false);
										clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

									}

								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Task) {
									MapiMessage Task = clientforexchange_input.fetchItem(msgInfo.getUniqueUri());
									MapiTask mapiTask = (MapiTask) Task.toMapiMessageItem();

									Date reciveddate = mapiTask.getStartDate();
									if (task_box.isSelected()) {
										if (reciveddate.after(taskfilterstartdate)
												&& reciveddate.before(taskfilterenddate)) {
											clientforexchange_output.createTask(subfolder, mapiTask);
											count_destination++;
										} else if (reciveddate.equals(taskfilterstartdate)
												|| reciveddate.equals(taskfilterenddate)) {

											clientforexchange_output.createTask(subfolder, mapiTask);
											count_destination++;
										}

									} else {

										clientforexchange_output.createTask(subfolder, mapiTask);
										count_destination++;

									}

								}

								else {
									MailMessage msg1 = clientforexchange_input.fetchMessage(msgInfo.getUniqueUri());
									Date reciveddate = msg1.getDate();
									String messageid = mailexchange(msg1, reciveddate, clientforexchange_output,
											subfolder);

									if (msgInfo.isRead()) {

										clientforexchange_output.setReadFlag(messageid, true);

									} else {
										clientforexchange_output.setReadFlag(messageid, false);
										clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

									}

								}
								lbl_progressreport.setText("Total Message Saved : " + count_destination + "    "
										+ folderInfo.getDisplayName() + " Extarcting messsage " + msgInfo.getSubject());
								if ((i % 100) == 0) {
									System.gc();
								}

							} catch (Exception e) {
								System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());

								e.printStackTrace();
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
								mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());
								continue;
							}

						}
					}
				}
				if (folderInfo.getChildFolderCount() > 0) {

					Convertexchange_subfolder_exchangeoutput(folderInfo, subfolder);

				}
			} catch (

			Exception e) {

				if (e.getMessage()
						.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")
						|| e.getMessage()
								.contains("No connection could be made because the target machine actively refused it.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
								"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
					Progressbar.setVisible(false);
					connectionHandle(e.getMessage());

				}
				mf.logger.warning("Exception" + e.getMessage() + "Message" + System.lineSeparator());
				continue;
			}

		}

	}

	public void Convertexchange_subfolder_exchangeoutput(ExchangeFolderInfo folderInfo, String Subfolder1) {

		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(folderInfo);

		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {
				ExchangeFolderInfo folderInfo1 = folderInfoCollection.get_Item(j);

//				if (stop) {
//				break;
//			}
				if (stop) {
					break;
				}
				Folder = folderInfo1.getDisplayName().replaceAll("[\\[\\]]", "");

				String Folder = getRidOfIllegalFileNameCharacters(folderInfo1.getDisplayName());

				String subfolder = Subfolder1;

				if (pstfolderlist.size() == countcheck) {
					break;
				}
				lbl_progressreport.setText("Getting Folder please wait ");

				for (int l = 0; l < pstfolderlist.size(); l++) {
//					if (stop) {
//					break;
//				}
					if (stop) {
						break;
					}
					if (listExchangdinal.get(l).getUri().equalsIgnoreCase(folderInfo1.getUri())) {
						countcheck++;

						Folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");
						listduplicacy.clear();

						System.out.println(folderInfo1.getFolderType());
						try {
							subfolder = clientforexchange_output
									.createFolder(Subfolder1, Folder, folderInfo1.getFolderType()).getUri();
						} catch (Exception e1) {
							subfolder = clientforexchange_output
									.createFolder(Subfolder1, Folder + 1, folderInfo1.getFolderType()).getUri();
						}

						ExchangeMessageInfoCollection msgCollection = clientforexchange_input
								.listMessages(folderInfo1.getUri());

						listduplicacy.clear();
						int i = 1;
						for (ExchangeMessageInfo msgInfo : msgCollection) {
							if (demo) {
								if (i == All_Data.demo_count) {
									break;
								}

							}
							i++;
//							if (stop) {
//							break;
//						}
							if (stop) {
								break;
							}
							if ((i % 100) == 0) {
								System.gc();
							}
							String strMessageURI = msgInfo.getUniqueUri();
							try {

								MailMessage msg1 = clientforexchange_input.fetchMessage(strMessageURI);

								MapiMessage msg = MapiMessage.fromMailMessage(msg1);
								if (msgInfo.getMessageInfoType() == MessageInfoType.Appointment) {
									Appointment appointment = clientforexchange_input
											.fetchAppointment(msgInfo.getUniqueUri());
									MailMessage mapi = new MailMessage();

									appointment.save(temppath + File.separator + calendarname(appointment) + ".ics",
											AppointmentSaveFormat.Ics);
									File file = new File(
											temppath + File.separator + calendarname(appointment) + ".ics");

									mapi.addAttachment(new Attachment(
											temppath + File.separator + calendarname(appointment) + ".ics"));
									file.delete();

									Date reciveddate = appointment.getStartDate();

									if (chckbx_calender_box.isSelected()) {
										if (reciveddate.after(Calenderfilterstartdate)
												&& reciveddate.before(Calenderfilterenddate)) {
											clientforexchange_output.appendMessage(subfolder, mapi);
											count_destination++;

										}

										else if (reciveddate.equals(Calenderfilterstartdate)
												|| reciveddate.equals(Calenderfilterenddate)) {

											clientforexchange_output.appendMessage(subfolder, mapi);
											count_destination++;
										}

									} else {
										clientforexchange_output.appendMessage(subfolder, mapi);
										count_destination++;

									}

									// System.out.println("Done");
								} else if (msgInfo.getMessageInfoType() == MessageInfoType.Contact) {
									Contact contact = clientforexchange_input.getContact(msgInfo.getUniqueUri());

									contact.save(temppath + File.separator + contactname(contact) + ".vcf",
											ContactSaveFormat.VCard);

									MailMessage mapi = new MailMessage();
									File file = new File(temppath + File.separator + contactname(contact) + ".vcf");
									mapi.addAttachment(
											new Attachment(temppath + File.separator + contactname(contact) + ".vcf"));
									file.delete();

									clientforexchange_output.appendMessage(subfolder, mapi);

									count_destination++;

								} else if (msg.getMessageClass().equals("IPM.StickyNote")
										|| msg.getMessageClass().equals("IPM.Task")) {
									try {

										lbl_progressreport
												.setText(Folder + " Extarcting messsage " + msg1.getSubject());
										// int messageprogress = (count * 100) / messagesize;
										// iprogressBar_message_p3.setValue(messageprogress);
										Date reciveddate = msg1.getDate();
										if (chckbx_Mail_Filter.isSelected()) {
											if (reciveddate.after(mailfilterstartdate)
													&& reciveddate.before(mailfilterenddate)) {
												clientforexchange_output.appendMessage(subfolder, msg1);
												count_destination++;

											}

										} else {
											clientforexchange_output.appendMessage(subfolder, msg1);
											count_destination++;

										}

									} catch (Exception e) {
										e.printStackTrace();
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
										mf.logger.warning(
												"Exception" + e.getMessage() + "Task" + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								}

								else {
									lbl_progressreport.setText(Folder + " Extarcting messsage " + msg.getSubject());
									// int messageprogress = (count * 100) / messagesize;
									// iprogressBar_message_p3.setValue(messageprogress);
									// count++;
									Date reciveddate = msg1.getDate();
									String messageid = mailexchange(msg1, reciveddate, clientforexchange_output,
											subfolder);

									if (msgInfo.isRead()) {

										clientforexchange_output.setReadFlag(messageid, true);

									} else {
										clientforexchange_output.setReadFlag(messageid, false);
										clientforexchange_input.setReadFlag(msgInfo.getUniqueUri(), false);

									}
								}
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

								mf.logger.warning(
										"Exception" + e.getMessage() + "Message" + " " + System.lineSeparator());
								continue;
							}

						}
					}
				}
				if (folderInfo1.getChildFolderCount() > 0) {

					Convertexchange_subfolder_exchangeoutput(folderInfo1, subfolder);

				}
			} catch (Exception e) {
				System.out.println("@@@@@@@@@@@@@@@ " + e.getMessage());
				if (e.getMessage()
						.contains("The operation 'FetchMessage' terminated. Timeout '100000' has been reached.")

						|| e.getMessage()
								.contains("No connection could be made because the target machine actively refused it.")
						|| e.getMessage().contains("ConnectFailure") || e.getMessage().contains("Rate limit hit")
						|| e.getMessage().contains("Operation failed")
						|| e.getMessage().contains("Operation has been canceled") || e.getMessage().contains(
								"The operation 'AppendMessage' terminated. Timeout '100000' has been reached")) {
					Progressbar.setVisible(false);
					connectionHandle(e.getMessage());

				} else {
					// System.out.println(">>>>>>>>>>>> " + e.getMessage());
				}
				continue;
			}

		}

	}

	public void Mapimessage_file(MapiMessage message) {
		String subname = mf.namingconventionmapi(message);
		try {
			File fu = null;
			if (filetype.equalsIgnoreCase("EML")) {
				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".eml");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".eml", SaveOptions.getDefaultEml());
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".eml",
								SaveOptions.getDefaultEml());

					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".eml",
							SaveOptions.getDefaultEml());
				}
			} else if (filetype.equalsIgnoreCase("HTML")) {

				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".html");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".html", SaveOptions.getDefaultHtml());
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".html",
								SaveOptions.getDefaultHtml());

					}

					new File(destination_path + File.separator + path + File.separator + "Attachment").mkdirs();

					for (int j = 0; j < message.getAttachments().size(); j++) {
						MapiAttachment att = message.getAttachments().get_Item(j);

						String s = getFileExtension(att.getDisplayName());
						String attFileName = getRidOfIllegalFileNameCharacters(
								att.getDisplayName().replace("." + s, ""));

						att.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + attFileName);

					}

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".html",
							SaveOptions.getDefaultHtml());
				}

			} else if (filetype.equalsIgnoreCase("MSG")) {

				try {
					fu = new File(destination_path + File.separator + path + File.separator + subname + ".msg");
					if (fu.isFile()) {
						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".msg", SaveOptions.getDefaultMsg());

					} else {

						message.save(destination_path + File.separator + path + File.separator + subname + ".msg",
								SaveOptions.getDefaultMsg());
					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".msg",
							SaveOptions.getDefaultMsg());
				}

			} else if (filetype.equalsIgnoreCase("EMLX")) {
				MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".emlx");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".emlx", SaveOptions.createSaveOptions(messagesavetype));
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
								SaveOptions.createSaveOptions(messagesavetype));

					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
							SaveOptions.createSaveOptions(messagesavetype));
				}
			} else if (filetype.equalsIgnoreCase("MHTML")) {

				try {
					fu = new File(destination_path + File.separator + path + File.separator + subname + ".mhtml");
					if (fu.isFile()) {
						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".mhtml", SaveOptions.getDefaultMhtml());

					} else {

						message.save(destination_path + File.separator + path + File.separator + subname + ".mhtml",
								SaveOptions.getDefaultMhtml());
					}
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".mhtml",
							SaveOptions.getDefaultMhtml());
				}

			}
			fu = null;

			count_destination++;

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Mailmessage_file(MailMessage message) {
		try {

			String subname = getRidOfIllegalFileNameCharacters(mf.namingconventionmail(message));

			File fu = null;
			if (filetype.equalsIgnoreCase("EML")) {
				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".eml");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".eml", SaveOptions.getDefaultEml());
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".eml",
								SaveOptions.getDefaultEml());

					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".eml",
							SaveOptions.getDefaultEml());
				}
			} else if (filetype.equalsIgnoreCase("HTML")) {

				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".html");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".html", SaveOptions.getDefaultHtml());
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".html",
								SaveOptions.getDefaultHtml());

					}
					new File(destination_path + File.separator + path + File.separator + "Attachment").mkdirs();

					for (int j = 0; j < message.getAttachments().size(); j++) {
						Attachment att = (Attachment) message.getAttachments().get_Item(j);

						String s = getFileExtension(att.getName());
						String attFileName = getRidOfIllegalFileNameCharacters(att.getName().replace("." + s, ""));

						att.save(destination_path + File.separator + path + File.separator + "Attachment"
								+ File.separator + attFileName);

					}

				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".html",
							SaveOptions.getDefaultHtml());
				}

			} else if (filetype.equalsIgnoreCase("MSG")) {

				try {
					fu = new File(destination_path + File.separator + path + File.separator + subname + ".msg");
					if (fu.isFile()) {
						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".msg", SaveOptions.getDefaultMsg());

					} else {

						message.save(destination_path + File.separator + path + File.separator + subname + ".msg",
								SaveOptions.getDefaultMsg());
					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".msg",
							SaveOptions.getDefaultMsg());
				}

			} else if (filetype.equalsIgnoreCase("EMLX")) {
				MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
				try {

					fu = new File(destination_path + File.separator + path + File.separator + subname + ".emlx");
					if (fu.isFile()) {

						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".emlx", SaveOptions.createSaveOptions(messagesavetype));
					} else {
						message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
								SaveOptions.createSaveOptions(messagesavetype));

					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(destination_path + File.separator + path + File.separator + subname + ".emlx",
							SaveOptions.createSaveOptions(messagesavetype));
				}
			} else if (filetype.equalsIgnoreCase("MHTML")) {

				try {
					fu = new File(destination_path + File.separator + path + File.separator + subname + ".mhtml");
					if (fu.isFile()) {
						message.save(destination_path + File.separator + path + File.separator + subname
								+ count_destination + ".mhtml", SaveOptions.getDefaultMhtml());

					} else {

						message.save(destination_path + File.separator + path + File.separator + subname + ".mhtml",
								SaveOptions.getDefaultMhtml());
					}
				} catch (Error e) {
					mf.logger.warning(
							"ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDate()
							+ System.lineSeparator());
					message.save(
							destination_path + File.separator + path + File.separator + +count_destination + ".mhtml",
							SaveOptions.getDefaultMhtml());
				}

			}
			fu = null;
			count_destination++;

		} catch (Exception e) {
			return;
		}

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
							model1.addRow(new Object[] { true, "", email, "Mailbox" });
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

	public boolean isValid(String email) {
		String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\." + "[a-zA-Z0-9_+&*-]+)*@" + "(?:[a-zA-Z0-9-]+\\.)+[a-z"
				+ "A-Z]{2,7}$";
		email = email.trim();
		Pattern pat = Pattern.compile(emailRegex);
		if (email == null)
			return false;
		return pat.matcher(email).matches();
	}

	public void Mapimess_CSV(MapiMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(mf.namingconventionmapi(message));

		try {
			String[] data1 = { message.getDeliveryTime().toString(), message.getSubject(), message.getBody(),
					message.getSenderEmailAddress(), message.getReplyTo(), message.getDisplayCc() };

			writer.writeNext(data1);

			if (message.getAttachments().size() > 0) {
				new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname).mkdirs();

			}

			// System.out.println("Getting Attachment");

			for (int j = 0; j < message.getAttachments().size(); j++) {

				MapiAttachment att = message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getLongFileName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getLongFileName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + File.separator + attFileName + "." + s);

			}

			count_destination++;

			// System.out.println("Done ");
		} catch (Error e) {
			mf.logger.warning(
					"ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime() + System.lineSeparator());
		}

		catch (Exception e) {
			mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
					+ System.lineSeparator());
		}

	}

	public void Mailmess_CSV(MailMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(mf.namingconventionmail(message));

		MapiMessage mp = MapiMessage.fromMailMessage(message);

		try {

			String[] data1 = { mp.getDeliveryTime().toString(), mp.getSubject(), mp.getBody(),
					mp.getSenderEmailAddress(), mp.getReplyTo(), mp.getDisplayCc() };

			writer.writeNext(data1);

			if (message.getAttachments().size() > 0) {
				new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname).mkdirs();

			}

			for (int j = 0; j < message.getAttachments().size(); j++) {
				Attachment att = (Attachment) message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + File.separator + attFileName + "." + s);

			}

			count_destination++;

		} catch (Error e) {
			mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
		}

		catch (Exception e) {
			mf.logger.warning(
					"Exception" + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());

		}

	}

	private void impersonationUser() {
		// clientforexchange_input.dispose();
		Contact[] contact = clientforexchange_input.getMailboxes();

		List<String> uscheck = new ArrayList<String>();

		DefaultTableModel dtm = (DefaultTableModel) table_userid.getModel();

		for (int j = 0; j < contact.length; j++) {
			EmailAddressList contactmailAdd = contact[j].getEmailAddresses();

			// clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2,
			// password_p2);

			for (EmailAddress emailAddress : contactmailAdd) {

				clientforexchange_input.impersonateUser(3, emailAddress.getAddress());
				try {
					ExchangeMailboxInfo mialinfo = clientforexchange_input.getMailboxInfo();
					clientforexchange_input.getMailboxInfo(mialinfo.getMailboxUri());
					String emailaddress = emailAddress.getAddress();
					System.out.println(emailaddress);
					// userList.add(emailaddress);
					if (!uscheck.contains(emailaddress)
							&& emailaddress.contains(usernameverification[usernameverification.length - 1])) {
						uscheck.add(emailaddress);
						// obThmailbox.ob.MessageLabel.setText(emailaddress);
						dtm.addRow(new Object[] { true, "", emailaddress, "Mailbox" });
					}
				} catch (Exception e) {

				}

			}

			try {
				clientforexchange_input.resetImpersonation();
				clientforexchange_input.dispose();

			} catch (Exception e) {

			}

		}

	}

	private static String getFileExtension(String fileName) {

		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	public ImapClient connectiontogmail_input() throws Exception {
		clientforimap_input = new ImapClient("imap.gmail.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoyahoo_input() throws Exception {
		clientforimap_input = new ImapClient("imap.mail.yahoo.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_input = clientforimap_input.createConnection();

		EmailClient.setSocketsLayerVersion2(true);

		// clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_input;
	}

	public ImapClient connectiontoaol_input() throws Exception {
		clientforimap_input = new ImapClient("imap.aol.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_input = clientforimap_input.createConnection();
		EmailClient.setSocketsLayerVersion2(true);

		// clientforimap_input.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_input;
	}

	public ImapClient connectiontoimap_input() throws Exception {
		clientforimap_input = new ImapClient(domain_p2, portmo, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		// clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoimap_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public IEWSClient conntiontooffice365_input() throws Exception {
		clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2, password_p2);
		clientforexchange_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforexchange_input;
	}

	public IEWSClient conntiontohotmail_input() throws Exception {
		clientforexchange_input = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p2,
				password_p2);
		clientforexchange_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);

		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforexchange_input;
	}

	public IEWSClient connectionwithexchangeserver_input() throws Exception {

		clientforexchange_input = EWSClient.getEWSClient("https://" + domain_p2 + "/ews/Exchange.asmx", username_p2,
				password_p2);
		clientforexchange_input.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_input;

	}

	public IEWSClient connectionwithexchangeserver_output() throws Exception {

		clientforexchange_output = EWSClient.getEWSClient("https://" + domain_p3 + "/ews/Exchange.asmx", username_p3,
				password_p3);
		clientforexchange_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;

	}

	public ImapClient connectiontogmail_output() throws Exception {
		clientforimap_output = new ImapClient("imap.gmail.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoYandex_output() throws Exception {
		clientforimap_output = new ImapClient("imap.yandex.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontozoho_output() throws Exception {
		clientforimap_output = new ImapClient("imap.zoho.in", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		// clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoinaws_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoaol_output() throws Exception {
		clientforimap_output = new ImapClient("imap.aol.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoyahoo_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.yahoo.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);

		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public IEWSClient conntiontooffice365_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	public IEWSClient conntiontohotmail_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p3,
				password_p3);

		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	private static String getRidOfIllegalFileNameCharacters(String strName) {
		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 80) {
			strLegalName = strLegalName.substring(0, 80);
		}
		return strLegalName;
	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {
				mf.logger.warning("Warning : " + e.getMessage());
			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {
				mf.logger.warning("Warning : " + e.getMessage());
			}
		}
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

	public void connectionHandle(String gotMessage) {
		lbl_progressreport.setText("INTERNET Connection  LOST ");
		label_12.setIcon(new ImageIcon(backup.class.getResource("/images.jpg")));

		while (true) {
			try {
				lbl_progressreport.setText("Connecting to Server Please Wait");
				conntiontooffice365_input();
				clientforexchange_input.impersonateUser(3, userid);
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();

				} else if (filetype.equalsIgnoreCase("Amazon Webmail")) {

					connectiontoinaws_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				label_12.setIcon(new ImageIcon(backup.class.getResource("/download.png")));
				lbl_progressreport.setText("Connection extablished Retriving Messasge");
				break;
			} catch (Exception e) {
				lbl_progressreport.setText("INTERNET Connection  LOST ");

			}

		}

		Progressbar.setVisible(true);

	}

	public void connectionHandle1(String gotMessage) {

		while (true) {
			try {

				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
					clientforexchange_input.impersonateUser(3, userid);
				}

				else if (filetype.equalsIgnoreCase("Amazon Webmail")) {

					connectiontoinaws_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				break;

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	String calendarname(Appointment ap) {
		String s = "";
		try {
			s = ap.getOrganizer().getDisplayName().replaceAll("[\\[\\]]", "");
		} catch (Exception e) {
			s = "Calendar";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	String contactname(Contact ap) {
		String s = "";
		try {
			s = ap.getDisplayName().replaceAll("[\\[\\]]", "");
		} catch (Exception e) {
			s = "Contact";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	void mapifile(MapiMessage message, Date reciveddate) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Mapimessage_file(message);
						foldermessagecount++;
					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Mapimessage_file(message);
						foldermessagecount++;
					}
				} else {
					Mapimessage_file(message);
					foldermessagecount++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Mapimessage_file(message);
					foldermessagecount++;
				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
					Mapimessage_file(message);
					foldermessagecount++;
				}
			} else {
				Mapimessage_file(message);
				foldermessagecount++;
			}
		}

	}

	void mapicsv(MapiMessage message, Date reciveddate, CSVWriter writer) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Mapimess_CSV(message, writer);
						foldermessagecount++;
					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Mapimess_CSV(message, writer);
						foldermessagecount++;
					}
				} else {
					Mapimess_CSV(message, writer);
					foldermessagecount++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Mapimess_CSV(message, writer);
					foldermessagecount++;
				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
					Mapimess_CSV(message, writer);
					foldermessagecount++;
				}
			} else {
				Mapimess_CSV(message, writer);
				foldermessagecount++;
			}
		}

	}

	String mapiexchange(MapiMessage message, Date reciveddate, String Folderuri) {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {

					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
				foldermessagecount++;
				count_destination++;

			}
		}
		return Messageid;
	}

	void mailfile(MailMessage message, Date reciveddate) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Mailmessage_file(message);

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Mailmessage_file(message);

					}
				} else {
					Mailmessage_file(message);

				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Mailmessage_file(message);

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
					Mailmessage_file(message);

				}
			} else {
				Mailmessage_file(message);

			}
		}

	}

	void mailmbox(MailMessage message, Date reciveddate, MboxrdStorageWriter wr) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						wr.writeMessage(message);
						count_destination++;
						foldermessagecount++;
					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						wr.writeMessage(message);
						count_destination++;
						foldermessagecount++;
					}
				} else {
					wr.writeMessage(message);
					count_destination++;
					foldermessagecount++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					wr.writeMessage(message);
					count_destination++;
					foldermessagecount++;
				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
					wr.writeMessage(message);
					count_destination++;
					foldermessagecount++;
				}
			} else {
				wr.writeMessage(message);
				count_destination++;
				foldermessagecount++;
			}
		}

	}

	void mailcsv(MailMessage message, Date reciveddate, CSVWriter writer) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Mailmess_CSV(message, writer);

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Mailmess_CSV(message, writer);

					}
				} else {
					Mailmess_CSV(message, writer);

				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Mailmess_CSV(message, writer);

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
					Mailmess_CSV(message, writer);

				}
			} else {
				Mailmess_CSV(message, writer);

			}
		}

	}

	String mailexchange(MailMessage message, Date reciveddate, IEWSClient clientforexchange_output1, String Folderuri)
			throws Exception {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

						count_destination++;

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

					count_destination++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

					count_destination++;

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {

					Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output1.appendMessage(Folderuri, message);

				count_destination++;

			}
		}
		return Messageid;
	}

	public ImapClient connectiontoGoDaddy_output() throws Exception {
		clientforimap_output = new ImapClient("imap.secureserver.net", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setConnectionCheckupPeriod(50000);

		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforimap_output;
	}

	public ImapClient connectiontonYandex_input() throws Exception {
		clientforimap_input = new ImapClient("imap.yandex.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_input = clientforimap_input.createConnection();

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforimap_input;
	}

	public ImapClient connectiontoHostgator_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforimap_output;
	}

	public ImapClient connectiontoicloud_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.me.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		iconnforimap_output = clientforimap_output.createConnection();

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforimap_output;
	}

	String mailimap(MailMessage message, Date reciveddate, String path) throws Exception {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = mf.duplicacymail(message);

			if (!listduplicacy.contains(input)) {
				System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						count_destination++;

					} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						count_destination++;

					}
				} else {

					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					count_destination++;
				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (reciveddate.after(mailfilterstartdate) && reciveddate.before(mailfilterenddate)) {
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					count_destination++;

				} else if (reciveddate.equals(mailfilterstartdate) || reciveddate.equals(mailfilterenddate)) {

					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					foldermessagecount++;
					count_destination++;
				}
			} else {
				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);

				count_destination++;

			}
		}
		return Messageid;
	}
}
