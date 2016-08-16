/**
 * @author zhangjie
 * @date 2014--8-30
 * @function  this function is used to convert MS Office documents and ODT 
 * Office documents to DHTML/PDF documents.
 * here, we use XDocReport based on Apache POI to implement this conversion, 
 * i.e, we use Apache POI to extract content from source MS/ODT Office 
 * documents, then SAX is used to generate DHTML documents, and iText is used
 * to generate PDF documents.
 * after testing for many formatted documents, it works well, please pay attention,
 * conversion to PDF document wlll take longer time to finish than conversioni to
 * DHTML, but the equality of conversion to PDF is better than conversion to
 * DHTML.
 */
package cn.edu.scu.util.word;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.pdf.ITextFontRegistry;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfWriter;
import com.sun.org.apache.commons.logging.Log;
import com.sun.org.apache.commons.logging.LogFactory;

import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;

public class Xdoc {

	private static Log log = LogFactory.getLog(Xdoc.class);
	private static String baseFontPath = null;

	public static String getBaseFontPath() {
		return Xdoc.baseFontPath;
	}

	public static void setBaseFontPath(String baseFontPath) {
		Xdoc.baseFontPath = baseFontPath;
	}

	/*
	 * public static void main(String args[]) {
	 * 
	 * System.out.println("the conversion to HTML  is begining ...");
	 * 
	 * docx2HTML(
	 * "/home/zhangjie/workspace/Foundation/src/cn/edu/scu/util/xdoc/川大教育基金会IMS--详细设计说明书.docx"
	 * ); docx2HTML("AdvancedTable.docx");
	 * docx2HTML("Docx4j_GettingStarted.docx");
	 * docx2HTML("FormattingTests.docx"); docx2HTML("HelloWorld.docx");
	 * docx2HTML("ooxml.docx"); docx2HTML("Resume.docx");
	 * docx2HTML("TableWithRowsColsSpan.docx");
	 * 
	 * System.out.println("the conversion to PDF  is begining ...");
	 * 
	 * docx2PDF("AdvancedTable.docx"); docx2PDF("Docx4j_GettingStarted.docx");
	 * docx2PDF("FormattingTests.docx"); docx2PDF("HelloWorld.docx");
	 * docx2PDF("ooxml.docx"); docx2PDF("Resume.docx");
	 * docx2PDF("TableWithRowsColsSpan.docx"); }
	 */
	/***
	 * 函数功能：转换2007版word为html
	 * @param docxPath  待转换的word绝对路径
	 * @return destPath  转换后的html绝对路径
	 * @throws Exception
	 */
	public static String docx2HTML(String docxPath) throws Exception {
		if (docxPath == null)
			return null;
		String docxDirname = docxPath.substring(0, docxPath.lastIndexOf("/"));
		String docxFilename = docxPath.substring(docxPath.lastIndexOf("/") + 1,
				docxPath.lastIndexOf("."));
		InputStream is = null;
		//这里创建html文件夹，不然会报文件不存在异常
		File directory = new File(docxDirname+"/html/");
		if(!directory.exists()){
			directory.mkdirs();
		}
		//这里设置html的保存路径
		String destPath = docxDirname + "/html/" + docxFilename+ ".html";
		try {
			long start = System.currentTimeMillis();
			is = new FileInputStream(docxPath);
			XWPFDocument document = new XWPFDocument(is);
			XHTMLOptions options = XHTMLOptions.create().indent(4);
			//img的src属性 后面会自动添加/word/media
			//这里就是设置img标签url的路径，这里必须设置相对路径，这里是./images/word/media/ + 图片名字	
			options.URIResolver(new BasicURIResolver("./images/"));								//13
			//这里设置 文件的保存路径    之后自动会添加 word\media子路径                                     					
			String htmlImagesPath = docxDirname + "/html/images/";                              //15
			FileImageExtractor extractor = new FileImageExtractor(new File(htmlImagesPath));    //16
			options.setExtractor(extractor);                                                    //17		
			XHTMLConverter.getInstance().convert(document,
					new FileOutputStream(destPath), options);
			log.info("Generate " + docxDirname + "/html/"
					+ docxFilename + ".html with "
					+ (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			is.close();
		}
		return destPath;
	}
/**
 * 
 * @param docxPath 绝对路径值
 * @return 对应的生成的pdf文件的绝对路径值
 */
	public static String docx2PDF(String docxPath) {
		if (docxPath == null)
			return null;
			
		String docxDirname = docxPath.substring(0, docxPath.lastIndexOf("/"));
		log.info("1号"+docxDirname);
		String docxFilename = docxPath.substring(docxPath.lastIndexOf("/") + 1,
				docxPath.lastIndexOf("."));
		log.info("2号"+docxFilename);
		String docxDirPath = docxDirname + "/pdf/"; 
		//File docxDir = new File(docxDirname + "/" + docxFilename + "/pdf/");
		File docxDir = new File(docxDirname + "/pdf/");
		if (!docxDir.exists())
			docxDir.mkdirs();

		try {
			long start = System.currentTimeMillis();

			// 1) Load DOCX into XWPFDocument
			InputStream is = new FileInputStream(new File(docxPath));
			XWPFDocument document = new XWPFDocument(is);

			// 2) Prepare Pdf options
			PdfOptions options = PdfOptions.create();

			options.fontProvider(new IFontProvider() {
				public Font getFont(String familyName, String encoding,
						float size, int style, Color color) {

					try {
						BaseFont bfChinese = BaseFont.createFont(
								Xdoc.baseFontPath, BaseFont.IDENTITY_H,
								BaseFont.EMBEDDED);
						Font fontChinese = new Font(bfChinese, size, style,
								color);
						if (familyName != null)
							fontChinese.setFamily(familyName);
						return fontChinese;
					} catch (Throwable e) {
						e.printStackTrace();
						return ITextFontRegistry.getRegistry().getFont(
								familyName, encoding, size, style, color);
					}
				}
			});

			// 3) Convert XWPFDocument to Pdf
//			OutputStream out = new FileOutputStream(new File(docxDirname + "/"
//					+ docxFilename + "/pdf/" + docxFilename + ".pdf"));
			OutputStream out = new FileOutputStream(new File(docxDirPath+docxFilename + ".pdf"));
			PdfConverter.getInstance().convert(document, out, options);

			log.info("Generate " + docxDirPath
					+ docxFilename + ".pdf with "
					+ (System.currentTimeMillis() - start) + "ms");
			//返回绝对路径值
//			return docxDirname + "/" + docxFilename + "/pdf/" + docxFilename
//					+ ".pdf";
			return docxDirPath + docxFilename + ".pdf";
		} catch (Throwable e) {
			e.printStackTrace();
			return null;
		}
	}
	/**
	 * 
	 * @param docxPath 绝对路径值
	 * @return 对应的生成的pdf文件的绝对路径值
	 */
		public static String jpg2PDF(String docxPath) {
			if (docxPath == null)
				return null;
				
			String docxDirname = docxPath.substring(0, docxPath.lastIndexOf("/"));
			log.info("1号"+docxDirname);
			String docxFilename = docxPath.substring(docxPath.lastIndexOf("/") + 1,
					docxPath.lastIndexOf("."));
			log.info("2号"+docxFilename);
			String docxDirPath = docxDirname + "/pdf/"; 
			File docxDir = new File(docxDirname + "/pdf/");
			if (!docxDir.exists())
				docxDir.mkdirs();
			
			try{//1：对图片文件通过TreeMap以名称进行自然排序
			  Map<Integer,File> mif = new TreeMap<Integer,File>();		  
			  //2：获取第一个Img的宽、高做为PDF文档标准
			  ByteArrayOutputStream baos = new ByteArrayOutputStream(2048*3);
			  InputStream is = new FileInputStream(docxPath);//mif.get(1));
			  for(int len;(len=is.read())!=-1;)
			    baos.write(len);
			  
			  baos.flush();
			  Image image = Image.getInstance(baos.toByteArray());
			  float width = image.getWidth();
			  float height = image.getHeight();
			  baos.close();
			  
			  //3:通过宽高 ，实例化PDF文档对象。
			  Document document = new Document(new Rectangle(width,height));
			  PdfWriter pdfWr = PdfWriter.getInstance(document, new FileOutputStream(docxDirPath+docxFilename + ".pdf"));
			    //4.1:读取到内存中
			    baos = new ByteArrayOutputStream(2048*3);
			    is = new FileInputStream(docxPath);
			    for(int len;(len=is.read())!=-1;)
			      baos.write(len);
			    baos.flush();
			    
			    //4.2通过byte字节生成IMG对象
			    image = Image.getInstance(baos.toByteArray());
			    Image.getInstance(baos.toByteArray());
			    image.setAbsolutePosition(0.0f, 0.0f);
			    
			    //4.3：添加到document中
			    document.open();
			   document.add(image);
			    
			    document.newPage();
			    baos.close();
			 // }
			  
			  //5：释放资源
			  document.close();
			  pdfWr.close();
			  return docxDirPath + docxFilename + ".pdf";
		} catch (Throwable e) {
			e.printStackTrace();
			return null;
		}
		}
}
