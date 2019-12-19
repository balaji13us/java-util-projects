package com.bs.java.utility.library.pdfmerger;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.util.*;
import java.util.List;

import javax.swing.*;
import java.io.*;
import java.awt.*;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

class Result {
	String message;
	boolean status;
}

public class PdgMergerSWT extends JFrame implements ActionListener {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private static List<String> pdfFileList;

	JLabel jLabelSourceFolderPath, jLabelDestinationFileName, jOutputLabel1, jOutputLabel2;
	JButton jButtonMergePDF;
	JTextField jTextFieldSourceFOlderPath, jTextFieldDestinationFile;
	JTextArea jOutputTextArea;
	JTable jt = new JTable();

	int sourcePathTextBoxX = 25, sourcePathTextBoxY = 20, destinationFileX = 25, destinationFileY = 60;
	int ButtonX = 25, ButtonY = 100, OutputX = 5, OutputY = 140;

	PdgMergerSWT(String SWTHeaderName) {
		super(SWTHeaderName);
		initialize();
	}

	public static void main(String[] args) throws IOException {
		new PdgMergerSWT("PDF Merger Utility");
	}

	public void initialize() {
		jLabelSourceFolderPath = new JLabel("Source Folder Path");
		jLabelSourceFolderPath.setBounds(sourcePathTextBoxX, sourcePathTextBoxY, 400, 30);
		add(jLabelSourceFolderPath);

		jTextFieldSourceFOlderPath = new JTextField("c:\\pdfsource");
		jTextFieldSourceFOlderPath.setBounds(sourcePathTextBoxX + 400, sourcePathTextBoxY, 400, 30);
		add(jTextFieldSourceFOlderPath);

		jLabelDestinationFileName = new JLabel("Provide Destination File Name with Complete Directory Path");
		jLabelDestinationFileName.setBounds(destinationFileX, destinationFileY, 400, 30);
		add(jLabelDestinationFileName);

		jTextFieldDestinationFile = new JTextField("c:\\pdfsource\\mergedOutput.pdf");
		jTextFieldDestinationFile.setBounds(destinationFileX + 400, destinationFileY, 400, 30);

		jButtonMergePDF = new JButton("Merge PDFs");
		jButtonMergePDF.addActionListener(this);
		jButtonMergePDF.setBounds(ButtonX, ButtonY, 200, 30);
		add(jButtonMergePDF);

		jOutputLabel1 = new JLabel("");
		jOutputLabel1.setBounds(OutputX, OutputY, 600, 150);
		add(jOutputLabel1);

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLayout(null);
		setLocation(100, 100);
		setVisible(true);
		setSize(1200, 350);

	}

	public static String convertToMultiline(String inputText) {
		return "<html>" + inputText.replaceAll("\n", "<br>") + "</html>";
	}

	public static Result getAllPDFFIles(String sourceFolderPath) {
		Result result = new Result();
		StringBuilder output = new StringBuilder("PDF Files found to be m erged \n");
		boolean pdfFound = false;
		File directory = new File(sourceFolderPath);
		if (null != directory) {
			File[] files = directory.listFiles();
			// If theis pathname does not denota a directoty, then listFiles() returns null
			if (null != files) {
				pdfFileList = new ArrayList<String>();
				for (File file : files) {
					if (file.isFile()) {
						System.out.println(file.getAbsolutePath());
						output.append(" ").append(file.getAbsolutePath()).append(" \n");
						if (file.getName().endsWith(".pdf")) {
							if (!pdfFound) {
								pdfFound = true;
							}
							pdfFileList.add(file.getName());
						}
					}
				}

			} else {
				System.out.println("No Files found");
				result.message = "Folder Path is not a Directory";
			}

			if (!pdfFound) {
				result.message = "No PDF FIles found to be merged";
				result.status = false;
			} else {
				result.message = output.toString();
				result.status = true;
			}
		}
		return result;
	}

	public static Result mergeAllPDFFIles(String sourceFolderPath, List<String> pdffileList, String DestinationFile) {
		File file = null;
		PDDocument doc = null;
		Result result = new Result();

		// Instantiating PDFMerger Utility Class
		PDFMergerUtility PDFMerger = new PDFMergerUtility();

		// setting the destination file
		PDFMerger.setDestinationFileName(DestinationFile);

		try {
			for (String fileName : pdfFileList) {
				file = new File(new StringBuilder().append(sourceFolderPath).append("\\").append(fileName).toString());
				doc = PDDocument.load(file);

				// adding the source files
				PDFMerger.addSource(file);

				// closing the documents
				doc.close();
			}

			// merging the documents
			PDFMerger.mergeDocuments(null);
			result.message = "PDF Documents merged successfull !!";
			result.status = true;
			System.out.println(result.message);

		} catch (Exception e) {
			e.printStackTrace();
			result.message = new StringBuilder("Exception Occurred: ").append(e.getClass().getSimpleName()).append(" ")
					.append(e.getMessage()).append(" ").append(e.getCause()).toString();
			result.status = false;
		}
		return result;

	}

	// @ Override
	public void actionPerformed(ActionEvent arg0) {
		StringBuilder output = new StringBuilder("Process Completed: ").append("\n");

		try {
			if (arg0.getSource() == jButtonMergePDF) {
				String sourceFolderPath = jTextFieldSourceFOlderPath.getText();
				String destinationFile = jTextFieldDestinationFile.getText();
				if (null == sourceFolderPath || sourceFolderPath.isEmpty()) {
					output.append("Folder Name cannot be empty").append("\n");

				} else if (null == destinationFile || destinationFile.isEmpty()) {
					output.append("Destination File name cannot be empty").append("\n");
				} else {
					jOutputLabel1.setText("");
					Result result = null;
					result = getAllPDFFIles(sourceFolderPath);
					System.out.println("sourceFolderPath : " + pdfFileList.toString());
					output.append("Files merged: \n").append(pdfFileList.toString()).append("\n");
					if (result.status) {
						result = mergeAllPDFFIles(sourceFolderPath, pdfFileList, destinationFile);
					}

					output.append(result.message).append("\n");
				}
				
				jOutputLabel1.setText(convertToMultiline(output.toString()));
				System.out.println("Merging Process Completed");
				
				output.append("Shrinking Process Started").append("\n");
				Thread.sleep(1000);
				Result resultOfShrink =  shrinkPDF(null, null);
				
				output.append(resultOfShrink.message).append("\n");
				
				output.append("Shrinking Process Completed").append("\n");
				jOutputLabel1.setText(convertToMultiline(output.toString()));
				System.out.println(output.toString());
			}
		} catch (Exception e) {
			e.printStackTrace();
			jOutputLabel1.setText(convertToMultiline("Exception Occured:"));
			JOptionPane
					.showMessageDialog(null,
							new JLabel(convertToMultiline(new StringBuilder("Exception Occurred: ")
									.append(e.getClass().getSimpleName()).append(" ").append(e.getMessage()).append(" ")
									.append(e.getCause()).toString())));
		}

	}

	public static Result shrinkPDF(String sourceFile, String destinationFile) {
		Result result = new Result();
		try {

			PDDocument pdDocument = new PDDocument();
			PDDocument oDocument = PDDocument.load(new File("c:\\pdfsource\\mergedOutput.pdf"));
			PDFRenderer pdfRenderer = new PDFRenderer(oDocument);
			int numberOfPages = oDocument.getNumberOfPages();
			PDPage page = null;

			for (int i = 0; i < numberOfPages; i++) {
				page = new PDPage(PDRectangle.LETTER);
				BufferedImage bim = pdfRenderer.renderImageWithDPI(i, 200, ImageType.RGB);
				PDImageXObject pdImage = JPEGFactory.createFromImage(pdDocument, bim);
				PDPageContentStream contentStream = new PDPageContentStream(pdDocument, page);
				float newHeight = PDRectangle.LETTER.getHeight();
				float newWidth = PDRectangle.LETTER.getWidth();
				contentStream.drawImage(pdImage, 0, 0, newWidth, newHeight);
				contentStream.close();

				pdDocument.addPage(page);
			}

			pdDocument.save("c:\\pdfsource\\compressedOutput.pdf");
			pdDocument.close();
			result.message = "PDF DOcuments merged successfull !!";
			result.status = true;
			System.out.println(result.message);

		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;

	}

}
