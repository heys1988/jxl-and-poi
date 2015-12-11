package heys;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JTextArea;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MatchThread implements Runnable {

	private JTextArea taInfoArea = null;
	private JProgressBar pbProg = null;
	private String daiYongPath;
	private String newFilePath;
	public void setTextArea(JTextArea taText)
	{
		taInfoArea = taText;
	}
	public void setProgressBar(JProgressBar pbBar)
	{
		pbProg = pbBar;
	}
	public void setDaiYongPath(String strPath)
	{
		daiYongPath = strPath;
	}
	@Override
	public void run() {
		// TODO Auto-generated method stub
		long tStart = System.currentTimeMillis();
		Workbook wbDaiYong = null;
		WritableWorkbook wbNewFile = null;
		File fDaiYong = null;
		String fDaiYongName = null;
		File fNewFile = null;
		String fNewFileName = null;
		try{
			fDaiYong = new File(daiYongPath);
			fDaiYongName = fDaiYong.getName();
			fNewFile = new File(newFilePath);
			fNewFileName = fNewFile.getName();
		}catch(NullPointerException e){
			e.printStackTrace();
			taInfoArea.append("空指针异常！" + "\n");
		}
		
		try {
			wbDaiYong = Workbook.getWorkbook(fDaiYong);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			taInfoArea.append("代用表" + "\""+fDaiYongName+"\"" +
					"操作异常！" + "\n");
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			taInfoArea.append("代用表" + "\""+fDaiYongName+"\"" +
					"操作异常！" + "\n");
			e.printStackTrace();
		}
		try {
			wbNewFile = Workbook.createWorkbook(fNewFile);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		WritableSheet shtNewFile = wbNewFile.createSheet("jxlsheet",0);
		Label label = new Label(0, 2, "A label record");//A3
		try {
			shtNewFile.addCell(label);
		} catch (RowsExceededException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (WriteException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		try {
			wbNewFile.write();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} 
		try {
			wbNewFile.close();
		} catch (WriteException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Sheet[] shtDaiYong = wbDaiYong.getSheets();
		if(1 != shtDaiYong.length)
		{
			taInfoArea.append("代用表" + "\""+fDaiYongName+"\"" +
					"中不止1个工作表！" + "\n");
			return;
		}

		int iDaiYongRows = shtDaiYong[0].getRows()-1;//忽略第1行字段
		List<String> list = new ArrayList<>();
		int count = 0;
		taInfoArea.append("正在读取代用表" + 
				"\""+fDaiYongName+"\"..." +"\n");
		int iOldIDColumn = 4;//第5列，即为第E列，从0开始数起
		for(int iR = 1; iR <= iDaiYongRows; iR++){
			++count;
			pbProg.setValue(count*100/iDaiYongRows);
			list.add(shtDaiYong[0].
					getCell(iOldIDColumn,iR).getContents());
		}
		wbDaiYong.close();
		taInfoArea.append("代用表读取完成。\n"+"\n");
		pbProg.setValue(0);
		int listsize = list.size();

		taInfoArea.append("代用表检查开始..."+"\n");
		
		count = 0;
		boolean bKeepOn = true;
		boolean bHint = false;
		int iDisplayRepeat = 10;
		for(int i = 0; i < listsize && bKeepOn; i++){
			String id1 = list.get(i);
			pbProg.setValue(i*100/(listsize-1));
			for(int j = i + 1; j < listsize && bKeepOn; j++){
				String id2 = list.get(j);
				if(id1.equals(id2)){
					if(false == bHint)
					{
						bHint = true;
						/*int iRet = JOptionPane.showConfirmDialog(null,
								"代用表异常，有重复条目，是否继续比较？",
								"提示", JOptionPane.YES_NO_OPTION);
						String strRet = null;*/
						int iRet = 0;
						String strRet = (String) JOptionPane.showInputDialog(null, "请输入需要列举的最大重复条目数量(1-100)：\n",
								"提示", JOptionPane.INFORMATION_MESSAGE,
								null, null,
								"10");
						try{									
							iRet = Integer.parseInt(strRet);
							if(iRet < 1 || iRet > 100)
							{
								throw new NumberFormatException();
							}
						}catch(NumberFormatException e){
							taInfoArea.append("\n操作不正确！"+"\n");
							bKeepOn = false;
							taInfoArea.append("\n程序终止运行！"+"\n");
							continue;
						}
						if(iRet <= 0)
						{
							bKeepOn = false;
							taInfoArea.append("\n程序终止运行！"+"\n");
							continue;
						}
						else
						{
							iDisplayRepeat = iRet;
						}
					}
					count++;							
					if(count <= iDisplayRepeat)
					{
						taInfoArea.append(count+":第"+(i+2)+"行与第"+(j+2)+"行重复"+"\n");
						//滚动条指向最下方，便于阅读新收到的消息。
						taInfoArea.setCaretPosition(taInfoArea.getText().length());
					}
					else if(iDisplayRepeat+1 == count)
					{
						taInfoArea.append("超过"+iDisplayRepeat+"对重复条目，若干重复条目未列举！"+"\n");
						taInfoArea.setCaretPosition(taInfoArea.getText().length());
					}
					else
					{
						count = iDisplayRepeat + 2;
					}
				}						
			}
		}
		pbProg.setValue(0);
		if(true == bKeepOn)
		{
			taInfoArea.append("代用表检查结束。"+"\n");
			taInfoArea.setCaretPosition(taInfoArea.getText().length());
		}
		/*
		pbProg.setValue(0);
		int listsize = list.size();
			
		
		
		//System.out.println("listsize"+listsize);
			
		long tEnd = System.currentTimeMillis();		
		long tElapse = (tEnd - tStart) / 1000;*/
		
	}
	public void setNewPath(String strPath) {
		// TODO Auto-generated method stub
		newFilePath = strPath;
	}

}