package com.github.wyleedp.poi.worker;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * <pre>
 * [2023.11.03]
 *   * Hello World!! 워드파일 생성 예제 
 *   * 생성된 워드파일은 사용자 임시폴더의 년월일시분초_HelloWorld.docx 파일로 생성된다.
 *     - 경로 예) C:\Users\wyleedp\AppData\Local\Temp\20231106101051_HelloWorld.docx
 * </pre>
 */
public class WordHelloWorld {

	public void exec() {
		System.out.println("워드파일 생성 시작");
		
		FileOutputStream fos = null;
		XWPFDocument documentWord = null;
		
		try {
			String tmpHome = System.getProperties().getProperty("java.io.tmpdir");
			String wordFileName = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss") + "_HelloWorld.docx";
			String wordFilePath = tmpHome + wordFileName;
			
			documentWord = new XWPFDocument();
			
			XWPFParagraph XWPFParagraph = documentWord.createParagraph();
			XWPFParagraph.setAlignment(ParagraphAlignment.LEFT);
			
			XWPFRun helloRun = XWPFParagraph.createRun();
			helloRun.setFontFamily("맑은 고딕");
			helloRun.setColor("2FB2F3");
			helloRun.setFontSize(20);
			helloRun.setText("Hello World!!!");
			
			fos = new FileOutputStream(wordFilePath);
			documentWord.write(fos);
			
			System.out.println("워드파일 생성완료 : " + wordFilePath + " [" + FileUtils.byteCountToDisplaySize(new File(wordFilePath).length()) + "]");
		}catch(Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(fos);
			try {
				documentWord.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		
	}
	
	public static void main(String[] args) {
		WordHelloWorld wordHelloWorld = new WordHelloWorld();
		wordHelloWorld.exec();
	}
	
}
