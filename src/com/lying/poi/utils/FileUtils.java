package com.lying.poi.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.URLDecoder;

public class FileUtils {
	private static final int SIZE = 2048;

	public static String getPath() {
		URL url = FileUtils.class.getProtectionDomain().getCodeSource().getLocation();
		String filePath = null;
		try {
			filePath = URLDecoder.decode(url.getPath(), "utf-8");// ת��Ϊutf-8����
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (filePath.endsWith(".jar")) {// ��ִ��jar�����еĽ�������".jar"
			// ��ȡ·���е�jar����
			filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);
		}
		
		if (filePath.endsWith(".exe")) {
			// ��ȡ·���е�jar����
			filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);
		}

		File file = new File(filePath);

		// /If this abstract pathname is already absolute, then the pathname
		// string is simply returned as if by the getPath method. If this
		// abstract pathname is the empty abstract pathname then the pathname
		// string of the current user directory, which is named by the system
		// property user.dir, is returned.
		filePath = file.getAbsolutePath();// �õ�windows�µ���ȷ·��
		return filePath;
	}

	public static String readFile(String filePath) {
		int len = 0;
		File file = new File(filePath);
		if (!file.exists()) {
			return "";
		}

		FileInputStream fis = null;
		StringBuffer sb = new StringBuffer();
		try {
			fis = new FileInputStream(file);
			byte[] buf = new byte[SIZE];
			while ((len = fis.read(buf)) != -1) {
				sb.append(new String(buf, 0, len));
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (fis != null) {
				// ����Դ
				try {
					fis.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

		return sb.toString();
	}
}
