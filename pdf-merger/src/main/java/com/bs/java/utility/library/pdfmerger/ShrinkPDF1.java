package com.bs.java.utility.library.pdfmerger;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
 
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
 
 
public class ShrinkPDF1 {
 
 public ShrinkPDF1() {
 
 }
 
 public static void main(String[] args) {
	 PDDocument oDocument = null;
  try {
 
   PDDocument pdDocument = new PDDocument();
   oDocument = PDDocument.load(new File("c:\\pdfsource\\mergedOutput.pdf"));
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
   
   System.out.println("PDF Compressed");
 
  } catch (IOException e) {
   e.printStackTrace();
  } finally {
	  try {
		oDocument.close();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  }
 
 
 }
 
}