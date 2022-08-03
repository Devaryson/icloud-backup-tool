package email.code;

import java.awt.Component;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

class Renderer extends DefaultTableCellRenderer {
	
	private static final long serialVersionUID = 1L;


@Override
	public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
			int row, int column) {
	
		return (Component) value;
	}

}