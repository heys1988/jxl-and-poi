package heys;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class XlsFileFilter extends FileFilter{

	@Override
	public boolean accept(File f) {
		// TODO Auto-generated method stub
		String nameString = f.getName();
		if(f.isDirectory()
				||nameString.toLowerCase().endsWith(".xls"))
		{
			return true;
		}
		return false;
	}

	@Override
	public String getDescription() {
		// TODO Auto-generated method stub
		return "Excel 文件(*.xls)";
	}

}