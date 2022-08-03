package email.code;

public class loadingThreadclassformailbox extends Thread {
	Main_Frame lf;
	loadingDialogformailbox ob;
	loadingThreadclassformailbox(Main_Frame lf)
	{
		this.lf=lf;
	}
	
	public void run()
	{
		ob=new loadingDialogformailbox(lf,true);
		visible();
	}
	
	void visible()
	{
		ob.setLocationRelativeTo(lf);
		ob.setVisible(true);
	}
	
	void close()
	{
		ob.setVisible(false);
	}
}
