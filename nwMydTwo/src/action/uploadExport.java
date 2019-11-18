package action;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * Servlet implementation class UploadExcel
 */
@WebServlet("/uploadExport")
public class uploadExport extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public uploadExport() {
        super();
        // TODO Auto-generated constructor stub
    }
	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		
		String  result =request.getSession().getAttribute("rest").toString(); 
		String  path=request.getSession().getAttribute("path").toString();
		InputStream in = new FileInputStream(path +File.separator+result + ".xlsx");  //下载的文件绝对路径
		// 获取下载文件的绝对路径
		OutputStream out = response.getOutputStream(); // 输出流进行编写
		String filename = new String(result.getBytes("iso-8859-1"), "utf-8");   //将下载的文件名称进行转码 utf-8
		// 下载的表格名称
		response.setHeader("Content-Type", getServletContext().getMimeType(filename+".xlsx"));
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("content-disposition", "attachment;filename=" + filename + ".xlsx"); // 设置content-disposition
		int len = 0;  			
		byte[] buffer = new byte[1024]; // 缓冲区
		while ((len = in.read(buffer)) != -1) { 
			out.write(buffer, 0, len); // 三参数,第一参数缓存区,第二是从什么位置开始,第三是大小为多少的(长度为多少)
		}
		in.close(); // 关闭
		out.flush();
		out.close();
	}
	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}
}
