package email.code;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;


import javax.swing.ImageIcon;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;

import javax.swing.JPanel;
import javax.swing.SwingConstants;
import javax.swing.border.EtchedBorder;
import java.awt.SystemColor;

public class loadingDialogformailbox extends JDialog {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	Main_Frame main;
	
	private final JPanel contentPanel = new JPanel();
	
	/**
	 * Launch the application.
	 */
	/*
	 * public static void main(String[] args) { try { JFrame parent = null; boolean
	 * modal = false; LoadingDialog dialog = new LoadingDialog(parent,modal);
	 * dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
	 * dialog.setVisible(true); } catch (Exception e) { e.printStackTrace(); } }
	 */
	
	/**
	 * Create the dialog.
	 * @param modal 
	 * @param parent 
	 */
	JLabel MessageLabel;
	public loadingDialogformailbox(JFrame parent, boolean modal) {
		super(parent,true);
		main=(Main_Frame) parent;
		setUndecorated(true);
		setResizable(false);
		setBounds(100, 100, 410, 197);
		getContentPane().setLayout(new BorderLayout());
		contentPanel.setBackground(Color.WHITE);
		contentPanel.setBorder(new EtchedBorder(EtchedBorder.RAISED, SystemColor.textHighlight, null));
		getContentPane().add(contentPanel, BorderLayout.CENTER);
		contentPanel.setLayout(null);
		
		MessageLabel = new JLabel("Loading....");
		MessageLabel.setForeground(SystemColor.textHighlight);
		MessageLabel.setBackground(Color.WHITE);
		MessageLabel.setFont(new Font("Calibri", Font.BOLD, 14));
		MessageLabel.setHorizontalAlignment(SwingConstants.CENTER);
		MessageLabel.setBounds(7, 124, 395, 30);
		contentPanel.add(MessageLabel);
		

		{
			JLabel lblNewLabel = new JLabel("");
			lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
			lblNewLabel.setIcon(new ImageIcon(loadingDialogformailbox.class.getResource("/Spinner.gif")));
			lblNewLabel.setBounds(35, 9, 337, 110);
			contentPanel.add(lblNewLabel);
		}
	}
}
