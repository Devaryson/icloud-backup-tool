package email.design;

import java.awt.Component;

import javax.swing.ImageIcon;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;



public class ImageRenderer extends DefaultTableCellRenderer {
	  /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	JLabel lbl = new JLabel();
	 // ImageIcon userfolderIcon = new ImageIcon(MainFrame.class.getResource("/user fol icon.png"));
	   ImageIcon userfolderIcon = new ImageIcon(ImageRenderer.class.getResource("/User-icon.png"));
	  public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
	      boolean hasFocus, int row, int column) {
	    lbl.setText((String) value);
	    lbl.setIcon(userfolderIcon);
	    lbl.setBounds(0, 0, 100, 100);
	    return lbl;
	  }}