package pojo;

import java.util.List;

public class Read {
	public String proname;//项目名称
	public String type; //需求类型
	public String title;//需求标题
	public String precedence;//需求优先级
	public String status;//需求状态
	public	String xianzhuang;//需求现状
	public	String description;	//需求描述
	public	int    yujidate;	//预计工作量
	public	String   id;			//id  --对应需求评论表中的需求id
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getProname() {
		return proname;
	}
	public void setProname(String proname) {
		this.proname = proname;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getPrecedence() {
		return precedence;
	}
	public void setPrecedence(String precedence) {
		this.precedence = precedence;
	}
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	public String getXianzhuang() {
		return xianzhuang;
	}
	public void setXianzhuang(String xianzhuang) {
		this.xianzhuang = xianzhuang;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public int getYujidate() {
		return yujidate;
	}
	public void setYujidate(int yujidate) {
		this.yujidate = yujidate;
	}
}
