package email.code;

import java.awt.event.ActionEvent;

import javax.swing.Action;
import javax.swing.KeyStroke;
import javax.swing.text.JTextComponent;
import javax.swing.text.TextAction;

public class SelectAll extends TextAction{
	 /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	public SelectAll()
	    {
	        super("Select All");
	        putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control S"));
	    }
	    public void actionPerformed(ActionEvent e)
	    {
	        JTextComponent component = getFocusedComponent();
	        component.selectAll();
	        component.requestFocusInWindow();
	    }
}
