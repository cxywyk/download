package action;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.PreparedStatement;

import db.GetConnection;
import pojo.Read;

/**
 * Servlet implementation class GwUser
 */
@WebServlet("/getExport")
public class getExport extends HttpServlet {
	private static final long serialVersionUID = 1L;
	ResourceBundle resource = ResourceBundle.getBundle("config");
	private String fileDir = resource.getString("upload.dir");
	private String MaxSizeStr = resource.getString("MaxSize");
	Integer MaxSize = Integer.parseInt(MaxSizeStr);

	/**
	 * @see HttpServlet#HttpServlet()
	 */
	public getExport() {
		super();
		// TODO Auto-generated constructor stub
	}

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		doPost(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		// PrintWriter p = response.getWriter();
		String type = request.getParameter("type");
		String precedence = request.getParameter("precedence");
		String proname = request.getParameter("proname");
		try {
			doUpload(request, type, precedence, proname);
			response.sendRedirect("/nwMydTwo/uploadExport"); // 重定向下载文档
		} catch (ClassNotFoundException | SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// p.write("保存数据成功,可以进行下载"); //成功
		// p.close();
	}

	private void doUpload(HttpServletRequest request, String type, String precedence, String proname)
			throws ClassNotFoundException, SQLException, IOException {
		String thisname = "excel" + System.currentTimeMillis(); // 文件名称
		request.getSession().setAttribute("rest", thisname);
		String realPath = request.getSession().getServletContext().getRealPath("/upload"); // 文件路径
		request.getSession().setAttribute("path", realPath);
		File filepath = new File(realPath);
		if (!filepath.exists()) {
			filepath.mkdirs();
		}
		Connection con = (Connection) new GetConnection().getCon(); // 获取连接
		String sql = "select  item_项目名称,item_需求类型,item_需求标题,item_需求优先级  ,statelabel ,item_现状,item_需求描述 ,item_预计工作量,id     from `tlk_所有需求_创建需求`  where  1=1 ";
		Map<String, String> map = new HashMap<>();// 类型的map
		map.put("1", "开发");
		map.put("2", "需求");
		map.put("3", "测试");
		map.put("4", "Bug");
		map.put("5", "维护");
		Map<String, String> map2 = new HashMap<>();// 优先级的map
		map2.put("1", "急");
		map2.put("2", "高");
		map2.put("3", "中");
		map2.put("4", "低");
		if (!type.equals("0")) // type传过来的是数值,所以如果不等于0的话直接根据map的键来获取map的值
		{
			sql += " and  item_需求类型 like '" + map.get(type) + "'";
		}
		if (!precedence.equals("0")) {
			sql += " and item_需求优先级 like '" + map2.get(precedence) + "'";
		}
		if (!proname.equals("123123")) // 123123为所有项目的id 所以不需要条件,只需查询所有项目即可
		{
			sql += " and item_项目名称  =(select item_项目名称 from `tlk_软件项目_新建项目` where item_项目id like '" + proname + "')";
		}
		PreparedStatement stm = (PreparedStatement) con.prepareStatement(sql);// 执行sql语句预编译放注入
		ResultSet executeQuery = stm.executeQuery(); // 执行
		List<Read> list = new ArrayList<Read>(); // read类的集合不存储需求评论
		FileOutputStream file = new FileOutputStream(realPath + File.separator + thisname + ".xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(); // 状态工作簿
		XSSFSheet sheet = workbook.createSheet(); // 创建工作表单
		XSSFRow row = sheet.createRow(0); // 创建HSSFRow对象 （行）
		// 创建多个单元格充当列头
		XSSFCell rowCell = row.createCell(0); // 创建XSSFCell对象(单元格)
		rowCell.setCellValue("系统");
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell = row.createCell(1); // 创建XSSFCell对象(单元格)
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell.setCellValue("需求类别");
		rowCell = row.createCell(2); // 创建XSSFCell对象(单元格)
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell.setCellValue("需求标题");
		rowCell = row.createCell(3); // 创建XSSFCell对象(单元格)
		rowCell.setCellValue("紧急程度");
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell = row.createCell(4); // 创建XSSFCell对象(单元格)
		rowCell.setCellValue("需求描述");
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell = row.createCell(5); // 创建XSSFCell对象(单元格)
		rowCell.setCellValue("跟进内容");
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		rowCell = row.createCell(6); // 创建XSSFCell对象(单元格)
		rowCell.setCellValue("状态");
		fontStyle(workbook, rowCell);
		hw(sheet, row);
		// 赋值阶段
		while (executeQuery.next()) {
			Read read = new Read();
			read.setProname(executeQuery.getString(1));
			read.setType(executeQuery.getString(2));
			read.setTitle(executeQuery.getString(3));
			read.setPrecedence(executeQuery.getString(4));
			read.setStatus(executeQuery.getString(5));
			read.setXianzhuang(executeQuery.getString(6));
			read.setDescription(executeQuery.getString(7));
			read.setYujidate(executeQuery.getInt(8));
			read.setId(executeQuery.getString(9));
			list.add(read); // 添加内容
		}
		for (int i = 0; i < list.size(); i++) {
			for (int j = 0; j < 1; j++) {
				row = sheet.createRow(i + 1);// 每次都在+1
				rowCell = row.createCell(0);
				rowCell.setCellValue(list.get(i).getProname()); // 系统
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(1);
				rowCell.setCellValue(list.get(i).getType()); // 类型
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(2);
				rowCell.setCellValue(list.get(i).getTitle()); // 标题
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(3);
				rowCell.setCellValue(list.get(i).getPrecedence()); // 紧急度
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(4);
				rowCell.setCellValue("1.现状:" + Filter(list.get(i).getXianzhuang()) + "\n2.需求:"
						+ Filter(list.get(i).getDescription()) + "\n3.计划时间:" + list.get(i).getYujidate() + "(人天)"); // 需求描述
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(5);
				// 取评论阶段
				String newsql = "select item_评论内容 from `tlk_所有需求_评论区`  where  item_需求id like '" + list.get(i).getId()
						+ "'"; // 通过需求id取评论
				stm = (PreparedStatement) con.prepareStatement(newsql);// 执行sql语句预编译放注入
				executeQuery = stm.executeQuery(); // 执行
				StringBuffer stu = new StringBuffer();
				int number = 1;
				while (executeQuery.next()) {
					stu.append(number + "." + executeQuery.getString(1) + "\n");
					number++;
				}
				// 赋值评论
				rowCell.setCellValue(stu.toString().trim()); // 需求跟进
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
				rowCell = row.createCell(6);
				rowCell.setCellValue(list.get(i).getStatus()); // 需求状态
				fontStyleCenter(workbook, rowCell);
				hwCenter(row);
			}
		}
		new GetConnection().closeCon();
		workbook.write(file);
		file.close();
	}

	/**
	 * 过滤区域将现状和需求的html代码过滤
	 * 
	 * @param name
	 * @return
	 */
	public String Filter(String name) {
		char[] abc = new char[] { 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q',
				'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '<', '/', '>', '&', '=', '"', '-', '0', '#', '(', ';', '1',
				'2', '3', '4', '5', '6', '7', '8', '9', ')', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
				'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };// 非中文字符
		StringBuffer buf = new StringBuffer(); // 进行内容获取并追加
		if (name != null && name != "") {
			// 替换区域
			String replace = name.replaceAll("<", "").replace("/", "").replace(">", "");
			// 进行过滤
			for (int w = 0; w < replace.length(); w++) {
				for (int e = 0; e < abc.length; e++) {
					if (replace.charAt(w) == abc[e]) {
						String replace2 = replace.replace(replace.charAt(w), 'n');
						replace = replace2.replace("微软雅黑", ""); // 赋值
					}
				}
			}
			for (int i = 0; i < replace.length(); i++) {
				if (replace.charAt(i) > 58) {
					buf.append(replace.charAt(i));
				}
			}
			return buf.toString().replace("n", "").trim(); // 返回的内容
		}
		return "";
	}

	/**
	 * 设置头部样式
	 * 
	 * @param workbook
	 * @param rowCell
	 */
	public void fontStyle(XSSFWorkbook workbook, XSSFCell rowCell) {
		XSSFCellStyle style = workbook.createCellStyle(); // 创建样式对象
		XSSFFont font = workbook.createFont(); // 创建font对象
		font.setFontName("宋体"); // 设置字体
		font.setBold(true); // 字体加粗
		font.setFontHeightInPoints((short) 12); // 字体大小
		style.setFont(font); // 设置字体
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平居中
		rowCell.setCellStyle(style); // 设置样式
	}

	/**
	 * 设置头部高度和所有宽度
	 */
	public void hw(XSSFSheet sheet, XSSFRow row) {
		sheet.setColumnWidth(0, 20 * 256); // 设置第1列宽度
		sheet.setColumnWidth(1, 10 * 256); // 设置第2列宽度
		sheet.setColumnWidth(2, 20 * 256); // 设置第3列宽度
		sheet.setColumnWidth(3, 10 * 256); // 设置第4列宽度
		sheet.setColumnWidth(4, 20 * 256); // 设置第5列宽度
		sheet.setColumnWidth(5, 20 * 256); // 设置第6列宽度
		sheet.setColumnWidth(6, 10 * 256); // 设置第7宽度
		row.setHeightInPoints(25); // 设置行高
	}

	/**
	 * 设置中间内容的样式
	 * 
	 * @param workbook
	 * @param rowCell
	 */
	public void fontStyleCenter(XSSFWorkbook workbook, XSSFCell rowCell) {
		XSSFCellStyle style = workbook.createCellStyle(); // 创建样式对象
		XSSFFont font = workbook.createFont(); // 创建font对象
		font.setFontName("宋体"); // 设置字体
		font.setFontHeightInPoints((short) 12); // 字体大小
		style.setFont(font); // 设置字体
		rowCell.setCellStyle(style); // 设置样式
	}

	/**
	 * 设置中间内容高度
	 */
	public void hwCenter(XSSFRow row) {
		row.setHeightInPoints(20); // 设置行高
	}
}
