package cn.edu.scu.util.word;

import javax.servlet.http.HttpServletRequest;

import org.jfree.util.Log;

/**
 * 根据前台传递的文件类型，分别将其转换为pdf文档，并将相应的路劲返回
 * 
 * @author Administrator
 * 
 */
public class ToPDF {
	/**
	 * 
	 * @param attachmentPath 文件相对路径值
	 * @param request
	 * @return 转换后的pdf相对路径值( upload/xx ...)
	 * @throws Exception
	 */
	public static String toPdf(String attachmentPath, HttpServletRequest request)
			throws Exception {
		//转换为绝对路径
		attachmentPath = request.getSession().getServletContext()
				.getRealPath("/")
				+ attachmentPath;
		Log.info("attachmenpath======="+attachmentPath);
		attachmentPath = attachmentPath.replaceAll("\\\\", "/");
		String previewFilename = "";
		int lastIndex = attachmentPath.lastIndexOf(".");
		// 得到文件四位后缀名
		String suffix = attachmentPath.substring(lastIndex);
		if (suffix.equals(".docx")) {
			//
			previewFilename = Xdoc.docx2PDF(attachmentPath);

		}if (suffix.equals(".jpg")) {
			//
			previewFilename = Xdoc.jpg2PDF(attachmentPath);

		} else if (suffix.equals(".pdf")) {
			previewFilename = attachmentPath;
		} else {
			previewFilename = doc.doc2Pdf(attachmentPath);
		}
		String atttachmentURL = previewFilename == null ? null
				: previewFilename.substring(
				previewFilename.indexOf("/upload/") + 1,
				previewFilename.length());
		return atttachmentURL;
	}
}
