package cn.edu.scu.util.word;

import java.io.File;
import java.io.IOException;

import javax.servlet.http.HttpServletRequest;

import cn.edu.scu.util.executeCmd.ExecuteCmd;

public class ToHTML {
	/***
	 * 
	 * @param attachmentPath   待转换word相对路径
	 * @param request
	 * @return atttachmentURL 转换后html的相对路径
	 * @throws Exception
	 */
	public static String toHtml(String attachmentPath, HttpServletRequest request)
			throws Exception {
		attachmentPath = request.getSession().getServletContext()
				.getRealPath("/")
				+ attachmentPath;
		attachmentPath = attachmentPath.replaceAll("\\\\", "/");
		String previewFilename = "";
		int lastIndex = attachmentPath.lastIndexOf(".");
		// 得到文件后缀名
		String suffix = attachmentPath.substring(lastIndex);
		if (suffix.equals(".docx")) {
			previewFilename = Xdoc.docx2HTML(attachmentPath);
		} else {
			System.out.print(attachmentPath);
			previewFilename = doc.doc2HTML(attachmentPath);
		}
		String atttachmentURL = previewFilename == null ? null
				: previewFilename.substring(
						previewFilename.indexOf("/upload/") + 1,
						previewFilename.length());
		return atttachmentURL;
	}
	/***
	 * 
	 * @param attachmentPath   待转换pdf相对路径(注：路径中不能有空格)
	 * @param request
	 * @return atttachmentURL 转换后html的相对路径
	 * @throws Exception
	 */
	public static String pdfToHtml(String attachmentPath, String htmlPath,HttpServletRequest request)
			throws Exception {
//		attachmentPath = request.getSession().getServletContext()
//				.getRealPath("/") + attachmentPath;
		attachmentPath = attachmentPath.replaceAll("\\\\", "/");
		String pdfDirname = attachmentPath.substring(0, attachmentPath.lastIndexOf("/"));
		String pdfFilename = attachmentPath.substring(attachmentPath.lastIndexOf("/") + 1,
				attachmentPath.lastIndexOf("."));
		int lastIndex = attachmentPath.lastIndexOf(".");
		String htmlDirName = pdfDirname + "/html/";
		File htmlDir = new File(htmlDirName);
		if(!htmlDir.exists()){
			htmlDir.mkdirs();
		}
		// 得到文件后缀名
		String suffix = attachmentPath.substring(lastIndex);
		StringBuilder sb= new StringBuilder();
		if (suffix.equals(".pdf")) {			 
			 sb.append("cmd /c d:/pdf2htmlEX/");
			 sb.append("pdf2htmlEX ");
			 sb.append(attachmentPath);
			 sb.append(" ");
			sb.append("--dest-dir ");
			 sb.append(htmlPath);			 
		}
		System.out.println(sb.toString());
		String attachmentURL = pdfDirname + "/html/" + pdfFilename+ ".html";
		attachmentURL = attachmentURL.substring(attachmentURL.indexOf("/upload/") + 1,
						attachmentURL.length());
		try {
            Process process = Runtime.getRuntime().exec(sb.toString());
            if (process.waitFor() == 0) {// 0 表示线程正常终止。
            	System.out.println("pdf->html转换完毕！");
            }else{
            	System.out.println("pdf->html转换失败！");
            	attachmentURL =  null;
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
		return attachmentURL;
	}
	
	/**
	 * 调用pdf2htmlEX将pdf文件转换为html文件,该方法处理了Runtime.exec()一直阻塞的情况，而上面方法没有
	 * @param command 调用exe的字符串
	 * @param pdfPath 需要转换的pdf文件名称 (注：路径中不能有空格)
	 * @param htmlDir 生成的html文件名称(注：html保存目录中路径分隔符必须是右斜杠“\”，左斜杠“/”不行)
	 * @return  转换成功则返回true，失败则返回false
	 */
	public static boolean pdf2html(String pdfPath,String htmlDir){
		htmlDir = htmlDir.replaceAll( "/","\\\\");
		String command = "cmd /c d:/project/pdf2htmlEX/pdf2htmlEX --dest-dir";
		command = command +" "+htmlDir+" "+pdfPath;
		System.out.println(command);
		if(ExecuteCmd.execute(command)){
			System.out.println("pdf->html转换完毕！");
			return true;
		}else{
			System.out.println("pdf->html转换失败！");
			return false;
		}
	}
	
  public static void main(String[] args) {
	  try {
		  String cmd = "cmd /c d:/pdf2htmlEX/pdf2htmlEX --dest-dir";
		  String htmldir = "d:\\apache-tomcat-7.0.55\\webapps\\Foundation\\upload\\reimbursement\\attachment\\20160601\\html";
		  String pdfPath = "d:\\apache-tomcat-7.0.55\\webapps\\Foundation\\upload\\reimbursement\\attachment\\20160601\\a5f79db0-27cd-4be8-9531-33de8a960ac9.pdf";
				  //"D:\\apache-tomcat-7.0.55\\webapps\\Foundation\\upload\\reimbursement\\attachment\\20160512\\";
		//pdfToHtml(dir+"75612a74-cf6b-4cce-a52d-689857adc5e3.pdf",dir+"html",null);
		  System.out.println(pdf2html(pdfPath,htmldir));
	} catch (Exception e) {
		e.printStackTrace();
	} 
  }
}
