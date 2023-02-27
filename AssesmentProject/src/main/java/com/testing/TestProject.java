package com.testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class TestProject {

	public static void main(String[] args) throws IOException {
		Scanner scanner = new Scanner(System.in);
		System.out.print("Enter directory location: ");
		String directoryPath = scanner.nextLine();
		System.out.print("Enter search string: ");
		String searchString = scanner.nextLine();
		scanner.close();
		searchDirectory(directoryPath, searchString);
	}

	private static void searchDirectory(String directoryPath, String searchString) throws IOException {
	    File directory = new File(directoryPath);
	    
	    System.out.println("Number of files present in the directory "+ directory.listFiles().length );
	    if (!directory.isDirectory()) {
	        System.out.println(directoryPath + " is not a directory.");
	        return;
	    }
	    String[] searchWords = searchString.split(",");
	    for (String searchWord : searchWords) {
	    	System.out.println('\n'+ "Results for Keyword: "+searchWord);
	        for (File file : directory.listFiles()) {
	            if (file.isFile() && file.getName().endsWith(".docx")) {
	                int count = searchFile(file, searchWord.trim());
	                if (count > 0) {
	                    System.out.println("File Name is : " + file.getName() + "    " +  " Keyword : " +  searchWord.trim()    + "    " +   " No of Occurence : " + count + "    " +  " Directory : "+ directoryPath);
	                }else {
	                	System.out.println( searchWord.trim() + " keyword is not present in word document "+ file.getName() );
	                }
	            }
	        }
	    }
	}

	private static int searchFile(File file, String searchString) throws IOException {
		int count = 0;
		try (FileInputStream fis = new FileInputStream(file); XWPFDocument document = new XWPFDocument(fis)) // XWPFDocument it is class which represent microsoft word documents 
		{
			for (XWPFParagraph paragraph : document.getParagraphs()) {
				if (paragraph.getText().contains(searchString)) {
					count++;
				}
			}
		}
		return count;
	}
}