package com.github.wyleedp.poi.worker;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.filefilter.DirectoryFileFilter;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * <pre>
 * [2023.11.03]
 *   * OS사용자 홈디텍토리의 폴더명 목록을 워드파일로 생성하는 예제
 *   * 생성된 워드파일은 OS사용자 임시폴더의 년월일시분초_UserHome.docx 파일로 생성된다.
 *     - 경로 예) C:\Users\wyleedp\AppData\Local\Temp\20231106095505_UserHome.docx
 * </pre>
 */
public class WordUserHomeDirListCreate {

	public void exec() {
		System.out.println("워드파일 생성 시작");
		
		FileOutputStream fos = null;
		XWPFDocument document = null;
		
		try {
			String userHome = System.getProperties().getProperty("user.home");
			String tmpHome = System.getProperties().getProperty("java.io.tmpdir");
			File userHomeDir = new File(userHome);
			//File tmpDir = new File(tmpHome);
			
			String wordFileName = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss") + "_UserHome.docx";
			String wordFilePath = tmpHome + wordFileName;
			//File wordFile = new File(wordFilePath);
			
			System.out.println("폴더명 목록 - UserHome 경로 : " + userHome);
			String[] list = userHomeDir.list(DirectoryFileFilter.DIRECTORY);
			
			document = new XWPFDocument();
			
			for(String fileName : list) {
				System.out.println("폴더명 : " + fileName);
				XWPFParagraph XWPFParagraph = document.createParagraph();
				XWPFRun run = XWPFParagraph.createRun();
				run.setText(fileName);
			}
			
			fos = new FileOutputStream(wordFilePath);
			document.write(fos);
			
			System.out.println("워드파일 생성완료 : " + wordFilePath + " [" + FileUtils.byteCountToDisplaySize(new File(wordFilePath).length()) + "]");
		}catch(Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(fos);
			
			try {
				document.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		
	}
	
	public static void main(String[] args) {
		WordUserHomeDirListCreate poiWorkerWordCreate = new WordUserHomeDirListCreate();
		poiWorkerWordCreate.exec();
	}
	
}
