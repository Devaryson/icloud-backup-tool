package email.code;

import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Font;
import java.awt.SystemColor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JTextField;
import javax.swing.border.EtchedBorder;
import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;

public class sigin extends JDialog {
	private static final long serialVersionUID = 1L;
	Main_Frame mf;
	JRadioButton rdbtnLiveExchange;
	private final JPanel contentPanel = new JPanel();
	private JTextField textField;
	private JPasswordField passwordField;
	static String password_p3 = "";
	static String domain_p3 = "";
	static String username_p3 = "";
	static Boolean check;
	JCheckBox chckbxShowPassword;
	JLabel label;
	JRadioButton rdbtnOffice;
	JCheckBox chckbxIMAdmin;
	private JTextField textField_dominname;
	String fileoption;
	
	private final ButtonGroup buttonGroup = new ButtonGroup();

	public sigin(JFrame parent) {
		super(parent, true);
		mf = (Main_Frame) parent;
		fileoption = Main_Frame.fileoption;
		System.out.println(fileoption);

		setUndecorated(true);
		setResizable(false);
		setBounds(100, 100, 550, 420);
	
		getContentPane().setLayout(new BorderLayout());
		contentPanel.setBackground(Color.WHITE);
		contentPanel.setBorder(new EtchedBorder(EtchedBorder.RAISED, SystemColor.textHighlight, null));
		getContentPane().add(contentPanel, BorderLayout.CENTER);
		contentPanel.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBackground(Color.WHITE);
		panel.setBorder(new EtchedBorder(EtchedBorder.RAISED, Color.LIGHT_GRAY, null));
		panel.setBounds(10, 271, 534, 38);
		getContentPane().add(panel);
		panel.setLayout(null);
		
		
		JPopupMenu menu = mf.menu;
		Cursor cursor = new Cursor(Cursor.HAND_CURSOR);

		JButton button = new JButton("");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				setVisible(false);
				Main_Frame.password_p3 = password_p3;
				Main_Frame.username_p3 = username_p3;
				Main_Frame.domain_p3 = domain_p3;

				CardLayout card = (CardLayout) restore.Cardlayout.getLayout();
				card.show(restore.Cardlayout, "panel_2");
				dispose();
			}
		});
		button.setBounds(515, 0, 35, 24);
		button.setRolloverEnabled(false);
		button.setRequestFocusEnabled(false);
		button.setOpaque(false);
		button.setFocusable(false);
		button.setFocusTraversalKeysEnabled(false);
		button.setFocusPainted(false);
		button.setDefaultCapable(false);
		button.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button.setIcon(new ImageIcon(Main_Frame.class.getResource("/close-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				button.setIcon(new ImageIcon(Main_Frame.class.getResource("/close-btn.png")));
			}
		});

		button.setIcon(new ImageIcon(Main_Frame.class.getResource("/close-btn.png")));
		button.setContentAreaFilled(false);
		button.setBorderPainted(false);
		panel.add(button);

		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(Main_Frame.class.getResource("/signin-ofic-topbar.png")));
		lblNewLabel.setBounds(0, 2, 550, 46);
		panel.add(lblNewLabel);

		JButton btn_Signin = new JButton("");
		panel.add(btn_Signin);
		btn_Signin.setRolloverEnabled(false);
		btn_Signin.setRequestFocusEnabled(false);
		btn_Signin.setOpaque(false);
		btn_Signin.setFocusable(false);
		btn_Signin.setFocusTraversalKeysEnabled(false);
		btn_Signin.setFocusPainted(false);
		btn_Signin.setDefaultCapable(false);
		btn_Signin.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Signin.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Signin.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_Signin.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
		btn_Signin.setContentAreaFilled(false);
		btn_Signin.setBorderPainted(false);
		btn_Signin.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				chckbxShowPassword.setEnabled(false);
				passwordField.setEnabled(false);
				textField.setEditable(false);
				check=rdbtnLiveExchange.isSelected();
				try {
					username_p3 = textField.getText().replaceAll("//s", "");
					username_p3 = username_p3.trim();
				} catch (Exception a) {
					username_p3 = "";
				}
				try {
					password_p3 = new String(passwordField.getPassword());
					password_p3 = password_p3.trim();
				} catch (Exception a) {
					password_p3 = "";
				}
				try {
					domain_p3 = textField_dominname.getText().replaceAll("//s", "");
					domain_p3 = domain_p3.trim();
				} catch (Exception a) {
					domain_p3 = "";
				}

				if (username_p3.equalsIgnoreCase("") || password_p3.equalsIgnoreCase("")) {

					if (username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(mf, "User name and Password fields cann't be empty",
								Main_Frame.messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (username_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "User name field cann't be empty", Main_Frame.messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (password_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(mf, "Password field cann't be empty", Main_Frame.messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					}
					chckbxShowPassword.setEnabled(true);
					passwordField.setEnabled(true);

					textField.setEditable(true);

				}else if (check&&textField_dominname.getText().isEmpty()) {

					JOptionPane.showMessageDialog(mf, "Domain Name field cann't be empty", Main_Frame.messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxShowPassword.setEnabled(true);
					passwordField.setEnabled(true);
					textField.setEditable(true);

				} else if (!(mf.isValid(username_p3))) {

					JOptionPane.showMessageDialog(mf, "Please enter a valid username", Main_Frame.messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					chckbxShowPassword.setEnabled(true);
					passwordField.setEnabled(true);
					textField.setEditable(true);

				} else {
					chckbxIMAdmin.setEnabled(false);
					Thread th = new Thread(new Runnable() {

						@Override
						public void run() {
							label.setVisible(true);
							try {
								Main_Frame.password_p3 = password_p3;
								Main_Frame.username_p3 = username_p3;
								Main_Frame.domain_p3 = domain_p3;
								rdbtnLiveExchange.setEnabled(false);
								rdbtnOffice.setEnabled(false);
								restore.fileoptioncheck = check;
								if (check) {

									Main_Frame.mailboxUri = "https://" + domain_p3 + "/ews/Exchange.asmx";
								}else {
									Main_Frame.mailboxUri ="https://outlook.office365.com/EWS/Exchange.asmx";
								}

								mf.conntiontooffice365_output();

								setVisible(false);
							} catch (Error e) {
								mf.logger.warning("Error :" + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								chckbxIMAdmin.setEnabled(true);
								if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(mf, "Application specific password required",
											Main_Frame.messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(mf, "Connection not established",
											Main_Frame.messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}
								
								e.printStackTrace();
								chckbxShowPassword.setEnabled(true);
								passwordField.setEnabled(true);
								textField.setEditable(true);
								label.setVisible(false);
								rdbtnLiveExchange.setEnabled(true);
								rdbtnOffice.setEnabled(true);

							}
						}
					});
					th.start();
				}

			}
		});
		btn_Signin.setBounds(402, 247, 129, 38);

		JLabel label_1 = new JLabel("");
		label_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/obottom.png")));
		label_1.setBounds(0, 374, 550, 46);
		panel.add(label_1);

		chckbxIMAdmin = new JCheckBox("Assign Impersonation right");
		chckbxIMAdmin.setFont(new Font("Tahoma", Font.BOLD, 12));
		chckbxIMAdmin.setBackground(Color.WHITE);
		chckbxIMAdmin.setBounds(191, 253, 203, 23);
		panel.add(chckbxIMAdmin);

		label = new JLabel("");
		panel.add(label);
		label.setBounds(263, 213, 42, 29);
		label.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));

		JLabel lblDomainName = new JLabel("Domain Name");
		lblDomainName.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblDomainName.setBounds(25, 172, 97, 29);
		if (!fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			lblDomainName.setVisible(false);
		}
		panel.add(lblDomainName);

		passwordField = new JPasswordField();
		passwordField.setHorizontalAlignment(JTextField.CENTER);
		passwordField.setComponentPopupMenu(menu);
		panel.add(passwordField);
		passwordField.setFont(new Font("Tahoma", Font.BOLD, 13));
		passwordField.setBounds(142, 131, 270, 30);

		textField = new JTextField();
		textField.setHorizontalAlignment(JTextField.CENTER);
		panel.add(textField);
		textField.setComponentPopupMenu(menu);
		textField.setFont(new Font("Tahoma", Font.BOLD, 13));
		textField.setBounds(142, 90, 270, 30);
		textField.setColumns(10);

		textField_dominname = new JTextField();
		textField_dominname.setHorizontalAlignment(JTextField.CENTER);
		textField_dominname.setBounds(143, 172, 269, 30);
		panel.add(textField_dominname);
		textField_dominname.setComponentPopupMenu(menu);
		if (!fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
			textField_dominname.setVisible(false);
		}
		textField_dominname.setColumns(10);

		JLabel lblMoreHelp = new JLabel("More Help");
		lblMoreHelp.setForeground(Color.RED);

		lblMoreHelp.setCursor(cursor);
		lblMoreHelp.setFont(new Font("Tahoma", Font.PLAIN, 14));

		lblMoreHelp.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				mf.openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");

			}
		});

		panel.add(lblMoreHelp);
		lblMoreHelp.setBounds(424, 57, 63, 21);
		panel.add(lblMoreHelp);

		chckbxShowPassword = new JCheckBox("Show Password");
		panel.add(chckbxShowPassword);
		
		chckbxShowPassword.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxShowPassword.setBackground(Color.WHITE);
		chckbxShowPassword.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField.setEchoChar((char) 0);
				}

				else {
					passwordField.setEchoChar('*');
				}
			}
		});
		chckbxShowPassword.setBounds(424, 137, 120, 23);

		JLabel lblPassword = new JLabel("Password");
		panel.add(lblPassword);
		lblPassword.setBounds(25, 131, 78, 30);
		lblPassword.setFont(new Font("Tahoma", Font.BOLD, 13));

		JLabel lblUserName = new JLabel("User Name");
		panel.add(lblUserName);
		lblUserName.setBounds(25, 90, 90, 30);
		lblUserName.setFont(new Font("Tahoma", Font.BOLD, 13));

		rdbtnOffice = new JRadioButton("Office 365");
		rdbtnOffice.setSelected(true);

		buttonGroup.add(rdbtnOffice);
		rdbtnOffice.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (rdbtnOffice.isSelected()) {
					lblDomainName.setVisible(false);
					textField_dominname.setVisible(false);
					lblNewLabel.setIcon(new ImageIcon(Main_Frame.class.getResource("/signin-ofic-topbar.png")));

				}
			}
		});
		rdbtnOffice.setFont(new Font("Tahoma", Font.BOLD, 11));
		rdbtnOffice.setBackground(Color.WHITE);
		rdbtnOffice.setBounds(143, 55, 109, 23);
		panel.add(rdbtnOffice);

		rdbtnLiveExchange = new JRadioButton("Live Exchange ");
		buttonGroup.add(rdbtnLiveExchange);
		if (fileoption.contains("Office 365 Backup & Restore")) {
			rdbtnLiveExchange.setVisible(false);
			rdbtnOffice.setVisible(false);
		} else if (fileoption.contains("Exchange Backup & Restore")) {
			rdbtnLiveExchange.setVisible(false);
			rdbtnOffice.setVisible(false);
			rdbtnLiveExchange.setSelected(true);
			lblNewLabel.setIcon(new ImageIcon(Main_Frame.class.getResource("/signin-exch-topbar.png")));
		}
		rdbtnLiveExchange.setFont(new Font("Tahoma", Font.BOLD, 11));
		rdbtnLiveExchange.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (rdbtnLiveExchange.isSelected()) {
					lblDomainName.setVisible(true);
					lblNewLabel.setIcon(new ImageIcon(Main_Frame.class.getResource("/signin-exch-topbar.png")));
					textField_dominname.setVisible(true);
				}
			}
		});
		rdbtnLiveExchange.setBackground(Color.WHITE);
		rdbtnLiveExchange.setBounds(281, 55, 131, 23);
		panel.add(rdbtnLiveExchange);
		label.setVisible(false);

	}
}
