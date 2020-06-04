package org.sluggard.oot.dao;

import java.util.List;

import org.apache.ibatis.annotations.Select;
import org.sluggard.oot.bean.TableInfo;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;

public interface SimpleDao extends BaseMapper<String> {
	
	@Select("select A.Table_Name, c.comments t_comments, A.column_name,A.data_type,A.data_length,A.Data_Scale,A.nullable,A.Data_default,B.comments c_comments from user_tab_columns A,user_col_comments B, user_tab_comments c where A.Table_Name = B.Table_Name and A.Column_Name = B.Column_Name and c.table_name=a.table_name and (a.table_name='financial_payables' or a.table_name like 'FIN%')")
	List<TableInfo> runSimpleSql();

}
