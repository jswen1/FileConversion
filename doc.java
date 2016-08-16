package cn.edu.scu.util.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.ConnectException;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.jruby.util.Pack.Converter;
import org.w3c.dom.Document;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class doc {

	/**
	 * 将Office文档转换为PDF. 运行该函数需要用到OpenOffice, OpenOffice下载地址为
	 * http://www.openoffice.org/
	 * 
	 * <pre>
	 * 方法示例: 
	 * String sourcePath = "F:\\office\\source.doc"; 
	 * String destFile = "F:\\pdf\\dest.pdf"; 
	 * Converter.office2PDF(sourcePath, destFile);
	 * </pre>
	 * 
	 * @param sourceFile
	 *            源文件, 绝对路径. 可以是Office2003-2007全部格式的文档, Office2010的没测试. 包括.doc,
	 *            .docx, .xls, .xlsx, .ppt, .pptx等. 示例: F:\\office\\source.doc
	 * @param destFile
	 *            目标文件. 绝对路径. 示例: F:\\pdf\\dest.pdf
	 * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
	 *         则表示操作成功; 返回1, 则表示转换失败
	 */
	public static int office2PDF(String sourceFile, String destFile)
			throws FileNotFoundException {
		try {
			File inputFile = new File(sourceFile);
			if (!inputFile.exists()) {
				return -1;// 找不到源文件, 则返回-1
			}

			// 如果目标路径不存在, 则新建该路径
			File outputFile = new File(destFile);
			if (!outputFile.getParentFile().exists()) {
				outputFile.getParentFile().mkdirs();
			}
			//String OpenOffice_HOME = "D:\\OpenOffice 4";// 这里是OpenOffice的安装目录
			String OpenOffice_HOME = "C:\\Program Files (x86)\\OpenOffice 4";// 这里是OpenOffice的安装目录
			// 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
			if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
				OpenOffice_HOME += "\\";
			}
			// 启动OpenOffice的服务
			String command = OpenOffice_HOME
					+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
			Process pro = Runtime.getRuntime().exec(command);
			// connect to an OpenOffice.org instance running on port 8100
			OpenOfficeConnection connection = new SocketOpenOfficeConnection(
					"127.0.0.1", 8100);
			connection.connect();
			// convert
			DocumentConverter converter = new OpenOfficeDocumentConverter(
					connection);
			converter.convert(inputFile, outputFile);

			// close the connection
			connection.disconnect();
			// 关闭OpenOffice服务的进程
			pro.destroy();
			return 0;
		} catch (ConnectException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return 1;
	}

	public static String doc2Pdf(String docxPath) throws Exception {
		if (docxPath == null)
			return null;

		String docxDirname = docxPath.substring(0, docxPath.lastIndexOf("/"));

		String docxFilename = docxPath.substring(docxPath.lastIndexOf("/") + 1,
				docxPath.lastIndexOf("."));

		File docxDir = new File(docxDirname + "/pdf/");
		if (!docxDir.exists())
			docxDir.mkdirs();
		String destFile = docxDirname +  "/pdf/"
				+ docxFilename + ".pdf";
		office2PDF(docxPath, destFile);
		return destFile;
	}
	/**
	 * word2003
	 * 
	 * @param inFile
	 * @return
	 */
	public static String doc2HTML(String docPath) {
		File inFile = new File(docPath);
		String docDirname = docPath.substring(0, docPath.lastIndexOf("/"));
		String docFilename = docPath.substring(docPath.lastIndexOf("/") + 1,
				docPath.lastIndexOf("."));
		final String imagedir = docDirname +  "/html/images/";
		File dirF = new File(docDirname +  "/html/");
		if (!dirF.exists())
			dirF.mkdirs();
		File dirI= new File(imagedir);
		if (!dirI.exists() || !dirI.isDirectory()) {
			dirI.mkdir();
		}
		String destFile = docDirname  + "/html/" + docFilename + ".html";
		try {		
			OutputStream	out = new FileOutputStream(new File(destFile));
			HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(
					inFile));
			
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
					DocumentBuilderFactory.newInstance().newDocumentBuilder()
							.newDocument());
			wordToHtmlConverter.setPicturesManager(new PicturesManager() {
				public String savePicture(byte[] content,
						PictureType pictureType, String suggestedName,
						float widthInches, float heightInches) {
					//return imagedir + suggestedName;
					return "./images/"+suggestedName;
				}
			});
			wordToHtmlConverter.processDocument(wordDocument);
			List<Picture> pics = wordDocument.getPicturesTable()
					.getAllPictures();
			if (pics != null) {
				for (int i = 0; i < pics.size(); i++) {
					Picture pic = (Picture) pics.get(i);
					try {
						pic.writeImageContent(new FileOutputStream(imagedir
								+ pic.suggestFullFileName()));
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					}
				}
			}
			Document htmlDocument = wordToHtmlConverter.getDocument();

			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(out);

			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.transform(domSource, streamResult);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		return destFile;
	}
}
