package com.rickauer.officeutil;

import java.io.*;
import java.nio.channels.FileChannel;
import de.schlichtherle.io.ArchiveDetector;

public class OpenOfficeUtils {

	public static void main(String[] args) {
		String source = "C:/Users/noNameForM3/Test/ODS.ods";
		String destination = "C:/Users/noNameForM3/Test/ODS_replaced.ods";

		copyFile(source, destination);

		String content = readOpenOfficeContent(source);

		content = content.replace("${test}", "Das ist ein Test");

		writeOpenOfficeContent(destination, content);
	}

	@SuppressWarnings({ "resource", "null" })
	public static String readOpenOfficeContent(String filename) {
		Reader is = null;

		try {

			de.schlichtherle.io.File file = new de.schlichtherle.io.File(filename + "/content.xml",
					ArchiveDetector.ALL);
			char[] fileContent = new char[(int) file.length()];

			is = new de.schlichtherle.io.FileReader(file); 
			is.read(fileContent);

			return new String(fileContent);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			try {
				is.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
	}

	@SuppressWarnings({ "resource", "null" })
	public static void writeOpenOfficeContent(String filename, String content) {
		Writer os = null;

		try {
			de.schlichtherle.io.File file = new de.schlichtherle.io.File(filename + "/content.xml",
					ArchiveDetector.ALL);
			os = new de.schlichtherle.io.FileWriter(file);
			os.write(content);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			try {
				os.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
	}

	@SuppressWarnings({ "resource", "null" })
	public static void copyFile(String in, String out) {
		FileChannel inChannel = null;
		FileChannel outChannel = null;

		try {
			inChannel = new FileInputStream(new File(in)).getChannel();
			outChannel = new FileOutputStream(new File(out)).getChannel();
			inChannel.transferTo(0, inChannel.size(), outChannel);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			try {
				inChannel.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
			try {
				outChannel.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
	}
}
