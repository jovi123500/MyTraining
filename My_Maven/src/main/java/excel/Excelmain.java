package excel;

import java.io.IOException;

public class Excelmain 
{

	public static void main(String[] args) throws IOException
	{
		Excelcode ec=new Excelcode();
		/*String n=ec.readdata(0,0);
		System.out.println("data="+n);
		
		String m=ec.readdata(0,1);
		System.out.println("data="+m);
		
		String x=ec.readdata(1,0);
		System.out.println("data="+x);
		
		String y=ec.readdata(1,1);
		System.out.println("data="+y);*/
		
		for(int i=0;i<ec.rowsize();i++)
		{
			for(int j=0;j<2;j++)
			{
				String n=ec.readdata(i,j);
				System.out.println("data="+n);
			}
		}
	}

}
