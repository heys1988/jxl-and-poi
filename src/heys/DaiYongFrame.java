package heys;

import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.swing.Box;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JToolBar;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import jxl.CellType;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

public class DaiYongFrame extends JFrame{

	private class poiListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub
			String strFileInPath = tfBOM.getText();
			FileInputStream fisInFile = null;
			HSSFWorkbook wb	= null;		
			try {
				fisInFile = new FileInputStream(strFileInPath);
				wb = new HSSFWorkbook(fisInFile);
				fisInFile.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			FileOutputStream fosFileOut = null;
			String strFileOutPath = tfNewFile.getText();
			File fileOut = new File(strFileOutPath);
			if(fileOut.exists()){
				fileOut.delete();
			}
			try {
				fosFileOut = new FileOutputStream(strFileOutPath);
	            wb.write(fosFileOut);
	            fosFileOut.close();
	        } catch (FileNotFoundException ex) {
	            System.out.println(ex.getMessage());
	        } catch (IOException ex) {
	            System.out.println(ex.getMessage());
	        }

			
		}

	}
	private class testListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub
			String strSrc = tfBOM.getText();
			String strCopy = tfNewFile.getText();
			if(strCopy.equals("")){
				return;
			}
			File fSrc = new File(strSrc);
			File fCopy = new File(strCopy);
			if(fCopy.exists()){				
				String strMsg = fCopy.getName() + "已存在，是否替换？";
				int iRet = JOptionPane.showConfirmDialog(null, 
						strMsg, "替换文件", JOptionPane.YES_NO_OPTION);

				if(JOptionPane.YES_OPTION == iRet){
					fCopy.delete();
				}

			}
			Workbook wbRead = null;
			WritableWorkbook wwbCopy = null;
			
			try {
				wbRead = Workbook.getWorkbook(fSrc);
			} catch (BiffException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				wwbCopy = Workbook.createWorkbook(fCopy, wbRead);
				wbRead.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			try {
				wwbCopy.write();
				
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				wwbCopy.close();
			} catch (WriteException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
					
		}

	}
	private class thdPiPei implements Runnable {

		@Override
		public void run() {
			// TODO Auto-generated method stub
			String driverName ="sun.jdbc.odbc.JdbcOdbcDriver";
			/*用第二种方法的注意点：Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
			 * 其中的空格一定要跟ODBC数据源管理器中的驱动名称一模一样，
			 * 一个空格都不能少
			*/
			/*String dbURL="jdbc:odbc:driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=Excel文件的路径";
			String dbURL="jdbc:odbc:driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" +
					"DBQ=D:\\TEST\\新旧编码对照表匹配BOM .xls";*/
			String dbURL="jdbc:odbc:driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" +
					"DBQ="+daiYongPath;
			Connection conn = null;
			PreparedStatement psql;
			ResultSet res;
			try {
				Class.forName(driverName);
				conn = DriverManager.getConnection(dbURL, "", "");
				//String sql = "select * from [sheet1$A3:S255]";
				/*String sql = "select * from [Sheet1$]" +
						"WHERE 更改前 = ?";
				psql = conn.prepareStatement(sql);
				psql.setString(1, "CPZ83054KP11G");//line 16656
				res = psql.executeQuery();
				while(res.next()){
					String id = res.getString("更改前");
					String des = res.getString("规格型号");
					System.out.println("更改前："+id);
					System.out.println("规格型号："+des);
				}*/
					
				String sql = "select * from [Sheet1$]";
				psql = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_UPDATABLE);
				res = psql.executeQuery();
				int firstrow = res.getRow();
				res.last();
				int lastrow = res.getRow();
				long count1 = 0, count2 = 0;
				long tStart = System.currentTimeMillis();
				/*for(int i = firstrow + 1; i <= lastrow; i++){
					res.absolute(i);
					String id1 = res.getString("更改前");
					count2++;
					//System.out.println("比较了"+count2+"次");
					System.out.println("第"+count2+"轮比较...");
					for(int j = i + 1; j <= lastrow; j++){
						res.absolute(j);
						String id2 = res.getString("更改前");
						boolean b = id1.equals(id2);
						if(b){
							count1++;
							//System.out.println("第"+i+"行与第"+j+"行重复！");
							if(0 == count1 % 50){
								System.out.println("第"+i+"行与第"+j+"行重复！");
							}
						}						
					}
				}*/
				res.first();
				//res.beforeFirst();
				//System.out.println(res.getString(5));
				//res.next();
				List<String> list = new ArrayList<>();
				list.add(res.getString("更改前"));
				int count = 1;
				taInfoArea.append("正在读取代用表..."+"\n");
				while(res.next()){
					++count;
					pbProg.setValue(count*100/lastrow);
					list.add(res.getString("更改前"));
				}
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
										"提示", JOptionPane.YES_NO_OPTION);*/
								String strRet = null;
								int iRet = 0;
								strRet = (String) JOptionPane.showInputDialog(null, "请输入需要列举的最大重复条目数量(1-100)：\n",
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
				
				//System.out.println("listsize"+listsize);
					
				long tEnd = System.currentTimeMillis();
				long tElapse = (tEnd - tStart) / 1000;
				//System.out.println("Time elapsed:"+tElapse+"sec");
				conn.close();				
			}catch(Exception e){
				e.printStackTrace();
			}finally{
				bnRun.setEnabled(true);			
			}

		}

	}
	
	private String daiYongPath;

	private String bomPath;

	private String newFilePath;
	private class runListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub
			daiYongPath = tfDaiYong.getText();
			bomPath = tfBOM.getText();
			newFilePath = tfNewFile.getText();
			boolean bRet = false;
			if(daiYongPath.equals("")){
				JOptionPane.showMessageDialog(null, "未打开代用表！", 
						"错误", JOptionPane.ERROR_MESSAGE);
				bRet = true;
			}
			else if(bomPath.equals(""))
			{
				JOptionPane.showMessageDialog(null, "未打开BOM表！", 
						"错误", JOptionPane.ERROR_MESSAGE);
				bRet = true;				
			}
			else if(newFilePath.equals(""))
			{
				JOptionPane.showMessageDialog(null, "未选择保存文件！", 
						"错误", JOptionPane.ERROR_MESSAGE);
				bRet = true;				
			}
			
			if(true == bRet)
			{				
				return;
			}
			taInfoArea.setText("");
			taInfoArea.append("Starting...\n\n");			
			//bnRun.removeActionListener(null);
			//bnRun.disable();
			bnRun.setEnabled(false);
			//mainProcess();
			//new Thread(new thdPiPei()).start();
			MatchThread thdMatch = new MatchThread();
			thdMatch.setTextArea(taInfoArea);
			thdMatch.setProgressBar(pbProg);
			thdMatch.setDaiYongPath(daiYongPath);
			thdMatch.setNewPath(newFilePath);
			new Thread(thdMatch).start();

		}

	}
	private class newFileListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub

		    JFileChooser fcOpenDlg = new JFileChooser();
		    fcOpenDlg.getCurrentDirectory();
		    fcOpenDlg.setCurrentDirectory(new File("D:/TEST"));
		    fcOpenDlg.setAcceptAllFileFilterUsed(false);
		    XlsFileFilter xlsFileFilter = new XlsFileFilter();
		    fcOpenDlg.addChoosableFileFilter(xlsFileFilter);

		    int index = fcOpenDlg.showDialog(null, "打开文件");
		    if (index == JFileChooser.APPROVE_OPTION) {
		    	String outFilePath = fcOpenDlg.getSelectedFile().getAbsolutePath();
		    	if(false == outFilePath.toLowerCase().endsWith(".xls"))
		    	{
		    		outFilePath += ".xls";
		    	}
		    	tfNewFile.setText(outFilePath);
		    }
		}

	}
	private class bomListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub

		    JFileChooser fcOpenDlg = new JFileChooser();
		    fcOpenDlg.getCurrentDirectory();
		    fcOpenDlg.setCurrentDirectory(new File("D:\\TEST"));
		    fcOpenDlg.setAcceptAllFileFilterUsed(false);
		    XlsFileFilter xlsFileFilter = new XlsFileFilter();
		    fcOpenDlg.addChoosableFileFilter(xlsFileFilter);

		    int index = fcOpenDlg.showDialog(null, "打开文件");
		    if (index == JFileChooser.APPROVE_OPTION) {
		    	//bomPath = fcOpenDlg.getSelectedFile().getAbsolutePath();
		    	tfBOM.setText(fcOpenDlg.getSelectedFile().getAbsolutePath());
		    }
		}

	}
	private class daiYongListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub

		    JFileChooser fcOpenDlg = new JFileChooser();
		    //设置默认的打开目录,如果不设的话按照window的默认目录(我的文档)
		    fcOpenDlg.getCurrentDirectory();
		    fcOpenDlg.setCurrentDirectory(new File(curPath));
		    fcOpenDlg.setCurrentDirectory(new File("D:\\TEST"));
		    //设置打开文件类型,此处设置成只能选择文件夹，不能选择文件
		    fcOpenDlg.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);//只能打开文件夹
		    //打开一个对话框
		    fcOpenDlg.setAcceptAllFileFilterUsed(false);
		    
		    XlsFileFilter xlsFileFilter = new XlsFileFilter();
		    fcOpenDlg.addChoosableFileFilter(xlsFileFilter);

		    int index = fcOpenDlg.showDialog(null, "打开文件");
		    if (index == JFileChooser.APPROVE_OPTION) {
		     //把获取到的文件的绝对路径显示在文本编辑框中
		     //daiYongPath = fcOpenDlg.getSelectedFile().getAbsolutePath();
		     tfDaiYong.setText(fcOpenDlg.getSelectedFile().getAbsolutePath());
		    }
		}

	}	
		
	/**
	 * @param args
	 */
	private static final long serialVersionUID = 1L;

	private static String curPath;
	
	private int iTFSize = 40;
	private int iBNSize = 150;
	private int iVertInterSpace = 20;

	private int iPreferredHeight = 50;

	private final Container cont;

	private final JPanel northPanel;

	private final JPanel northPanel2;

	private final JPanel northPanel3;

	private final JPanel northPanel4;

	private final JPanel southPanel;

	private final Box northBoxInPanel;

	private final JButton bnDaiYong;

	private final JTextField tfDaiYong;

	private final JButton bnBOM;

	private final JTextField tfBOM;

	private final JButton bnNewFile;

	private final JButton bnRun;

	private final JTextArea taInfoArea;

	private final JScrollPane scrollInfoArea;

	private final JToolBar tbStatus;

	private final JLabel lbStatus;

	private final JProgressBar pbProg;

	private final JTextField tfNewFile;

	private JButton bnTest;

	private JButton bnPOI;
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		curPath = System.getProperty("user.dir");
		DaiYongFrame frame = new DaiYongFrame();
		Integer dat = new Integer(-12 % 5);
		frame.taInfoArea.append(dat.toString());
		//frame.setVisible(true);
		//new DaiYongFrame();
	}

	private void mainProcess() {
		// TODO Auto-generated method stub		
	}

	public DaiYongFrame(){
		super();
		winInit();		
		cont = getContentPane();
		northBoxInPanel = Box.createVerticalBox();
		cont.add(northBoxInPanel, BorderLayout.NORTH);
		northPanel = new JPanel();
		northBoxInPanel.add(northPanel);
		northPanel2 = new JPanel();
		northBoxInPanel.add(northPanel2);
		northPanel3 = new JPanel();
		northBoxInPanel.add(northPanel3);
		northBoxInPanel.add(Box.createVerticalStrut(iVertInterSpace));
		northPanel4 = new JPanel();
		northBoxInPanel.add(northPanel4);
		northBoxInPanel.add(Box.createVerticalStrut(iVertInterSpace));
		
		
		bnDaiYong = new JButton("选择代用关系表");
		iPreferredHeight   = (int)bnDaiYong.getPreferredSize().getHeight();
		//bnDaiYong.setSize(iBNSize, (int)getPreferredSize().getHeight());
		bnDaiYong.setPreferredSize(new Dimension(iBNSize, iPreferredHeight));
		northPanel.add(bnDaiYong);
		tfDaiYong = new JTextField(iTFSize);
		northPanel.add(tfDaiYong);
		bnBOM = new JButton("选择要代用的BOM");
		bnBOM.setPreferredSize(new Dimension(iBNSize, iPreferredHeight));
		northPanel2.add(bnBOM);
		tfBOM = new JTextField(iTFSize);
		northPanel2.add(tfBOM);
		bnNewFile = new JButton("选择保存路径");
		bnNewFile.setPreferredSize(new Dimension(iBNSize, iPreferredHeight));
		northPanel3.add(bnNewFile);
		tfNewFile = new JTextField(iTFSize);
		northPanel3.add(tfNewFile);
		bnRun = new JButton("运行");
		northPanel4.add(bnRun);
		bnTest = new JButton("测试");
		northPanel4.add(bnTest);
		bnPOI = new JButton("POI");
		northPanel4.add(bnPOI);
		taInfoArea = new JTextArea();
		Font ftArea = new Font(null, Font.PLAIN, 16);
		taInfoArea.setFont(ftArea);
		//taInfoArea.setPreferredSize(new Dimension(60,60));
		taInfoArea.setTabSize(2);//按下Tab键的间隔
		taInfoArea.setLineWrap(true);
		taInfoArea.setEditable(false);
		//taInfoArea.setWrapStyleWord(true);
		scrollInfoArea = new JScrollPane(taInfoArea); 
		//scrollInfoArea.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); 
		scrollInfoArea.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); 
		cont.add(scrollInfoArea, BorderLayout.CENTER);
		southPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
		cont.add(southPanel, BorderLayout.SOUTH);
		pbProg = new JProgressBar();
		southPanel.add(pbProg);
		pbProg.setStringPainted(true);
		tbStatus = new JToolBar();
		southPanel.add(tbStatus);
		tbStatus.setFloatable(false);
		String strLabel = "深圳市杰科电子有限公司 DTV研发部 何越盛 2015-8-19";
		lbStatus = new JLabel(strLabel);
		tbStatus.add(lbStatus);
		//new Thread(new MyThread(cont, southPanel, pbProg)).start();
		
		
		bnDaiYong.addActionListener(new daiYongListener());
		bnBOM.addActionListener(new bomListener());
		bnNewFile.addActionListener(new newFileListener());
		bnRun.addActionListener(new runListener());
		bnTest.addActionListener(new testListener());
		bnPOI.addActionListener(new poiListener());
		
	}
	public void winInit() {
		// TODO Auto-generated method stub
		
		int iWidth = 800;
		int iHeight = iWidth * 3 / 4;
		setSize(iWidth, iHeight);
		Dimension screensize = Toolkit.getDefaultToolkit().getScreenSize(); 
		Dimension framesize = getSize(); 
		int x = (int)screensize.getWidth()/2 - (int)framesize.getWidth()/2; 
		int y = (int)screensize.getHeight()/2 - (int)framesize.getHeight()/2; 
		setLocation(x,y);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setResizable(false);
		setLayout(new BorderLayout());
		setVisible(true);
	}
}