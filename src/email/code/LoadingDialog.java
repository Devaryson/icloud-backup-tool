package email.code;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.SwingConstants;
import javax.swing.border.EtchedBorder;
import java.awt.SystemColor;

public class LoadingDialog extends JDialog {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	Main_Frame main;
	JButton stopButton ;
	private final JPanel contentPanel = new JPanel();
	
	
	JLabel MessageLabel;
	public LoadingDialog(JFrame parent, boolean modal) {
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
		
		stopButton = new JButton("");
		stopButton.setBorderPainted(false);
		stopButton.setContentAreaFilled(false);
		stopButton.setIcon(new ImageIcon(LoadingDialog.class.getResource("/stop-btn.png")));
		stopButton.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				stopButton.setIcon(new ImageIcon(LoadingDialog.class.getResource("/stop-hvr-btn.png")));
			}
			@Override
			public void mouseExited(MouseEvent e) {
				stopButton.setIcon(new ImageIcon(LoadingDialog.class.getResource("/stop-btn.png")));
			}
		});
		stopButton.setFocusPainted(false);
		stopButton.setFont(new Font("Calibri", Font.BOLD, 16));
		stopButton.setBackground(Color.WHITE);
		stopButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				int result = JOptionPane.showConfirmDialog(LoadingDialog.this,"Are you sure, You want to stop the Process?",Main_Frame.messageboxtitle,
						JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE,new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if(result==JOptionPane.YES_OPTION)
				{
					main.stop_tree=true;
					dispose();
				}
			}
		});
		stopButton.setBounds(292, 160, 110, 30);
		contentPanel.add(stopButton);
		{
			JLabel lblNewLabel = new JLabel("");
			lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
			lblNewLabel.setIcon(new ImageIcon(LoadingDialog.class.getResource("/Spinner.gif")));
			lblNewLabel.setBounds(35, 9, 337, 110);
			contentPanel.add(lblNewLabel);
		}
	}
}
