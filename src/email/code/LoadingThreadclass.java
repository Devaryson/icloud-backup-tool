package email.code;

public class LoadingThreadclass extends Thread {
	
	Main_Frame lf;
	LoadingDialog ob;
	LoadingThreadclass(Main_Frame lf)
	{
		this.lf=lf;
	}
	
	public void run()
	{
		ob=new LoadingDialog(lf,true);
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
