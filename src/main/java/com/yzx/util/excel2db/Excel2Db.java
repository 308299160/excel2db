package com.yzx.util.excel2db;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.metadata.Sheet;
import com.util.property.PropertyUtil;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.ServiceLoader;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel2Db {

    private Pattern tableNamePattern = Pattern.compile("(\\S+)(?:\\()(?<=\\()(\\w+)(?=\\))");

    private String currentTableName;
    private String tableComment;
    private List<FieldProperty> fieldProperties = new ArrayList<>();
    private Connection conn;

    private static final String FIELD_NAME = "字段";
    private int FIELD_NAME_INDEX;
    private static final String FIELD_TYPE = "类型";
    private int FIELD_TYPE_INDEX;
    private static final String PRIMARY_KEY = "主键";
    private int PRIMARY_KEY_INDEX;
    private static final String UNIQUE = "唯一";
    private int UNIQUE_INDEX;
    private static final String ALLOW_NULL = "是否可为空";
    private int ALLOW_NULL_INDEX;
    private static final String DEFAULT_VALUE = "默认值";
    private int DEFAULT_VALUE_INDEX;
    private static final String COMMENT = "备注";
    private int COMMENT_INDEX;

    private static final List<String> TITLES = new ArrayList<>(6);
    static {
        TITLES.add(FIELD_NAME);
        TITLES.add(FIELD_TYPE);
        TITLES.add(PRIMARY_KEY);
        TITLES.add(UNIQUE);
        TITLES.add(ALLOW_NULL);
        TITLES.add(DEFAULT_VALUE);
        TITLES.add(COMMENT);
    }

    private boolean checkTableIsNotEmpty() {
        return currentTableName != null && !currentTableName.equals("");
    }

    public void excel2Db(String jdbcUrl, String account, String password, String filePath) throws IOException, ClassNotFoundException, SQLException {
        ServiceLoader<Driver> drivers = ServiceLoader.load(Driver.class);
        Driver driver = drivers.iterator().next();
        conn = DriverManager.getConnection(jdbcUrl, account, password);

        InputStream inputStream = new FileInputStream(filePath);
        List<Object> data = EasyExcelFactory.read(inputStream, new Sheet(1));
        for(Object obj : data) {
            if(obj == null) {
                continue;
            }
            List<String> cells = (List<String>) obj;
            // 统计数量
            int realLength = 0;
            for(String c : cells) {
                if(c != null) {
                    realLength ++;
                }
            }

            if(realLength == 0) {
                continue;
            }
            if(realLength == 1) {
                // 表名
                Matcher matcher = tableNamePattern.matcher(cells.get(0));
                while(matcher.find()) {
                    String newTableComment = matcher.group(1);
                    String newCurrentTableName = matcher.group(2);
                    if(checkTableIsNotEmpty() && fieldProperties.size() > 0) {
                        createTable();
                    }

                    currentTableName = newCurrentTableName;
                    tableComment = newTableComment;
                    fieldProperties.clear();
                }

            }
            else {
                if(TITLES.contains(cells.get(0))){
                    FIELD_NAME_INDEX = cells.indexOf(FIELD_NAME);
                    FIELD_TYPE_INDEX = cells.indexOf(FIELD_TYPE);
                    PRIMARY_KEY_INDEX = cells.indexOf(PRIMARY_KEY);
                    UNIQUE_INDEX = cells.indexOf(UNIQUE);
                    ALLOW_NULL_INDEX = cells.indexOf(ALLOW_NULL);
                    DEFAULT_VALUE_INDEX = cells.indexOf(DEFAULT_VALUE);
                    COMMENT_INDEX = cells.indexOf(COMMENT);
                    continue;
                }
                if(checkTableIsNotEmpty()) {
                    //统计字段
                    FieldProperty property = new FieldProperty();
                    if (FIELD_NAME_INDEX > -1) {
                        property.setPropertyName(cells.get(FIELD_NAME_INDEX));
                    }
                    if (FIELD_TYPE_INDEX > -1) {
                        property.setType(cells.get(FIELD_TYPE_INDEX));
                    }
                    if (PRIMARY_KEY_INDEX > -1) {
                        property.setPrimaryKey(isY(cells.get(PRIMARY_KEY_INDEX)));
                    }
                    if (UNIQUE_INDEX > -1) {
                        property.setUnique(isY(cells.get(UNIQUE_INDEX)));
                    }
                    if (ALLOW_NULL_INDEX > -1) {
                        property.setAllowNull(isY(cells.get(ALLOW_NULL_INDEX)));
                    }
                    if (DEFAULT_VALUE_INDEX > -1) {
                        property.setDefaultValue(cells.get(DEFAULT_VALUE_INDEX));
                    }
                    if (COMMENT_INDEX > -1) {
                        property.setComment(cells.get(COMMENT_INDEX));
                    }
                    fieldProperties.add(property);
                }
            }


        }
        if(checkTableIsNotEmpty() && fieldProperties.size() > 0) {
            createTable();
        }
        inputStream.close();
        conn.close();
    }

    private boolean isY(String s) {
        return s != null && s.toUpperCase().equals("Y");
    }

    private void createTable() {
        // 开始建表
        System.out.println("开始建表--->"+currentTableName);

        StringBuilder primayKey = new StringBuilder();
        StringBuilder uniqueKey = new StringBuilder();

        StringBuilder sb = new StringBuilder();
        sb.append("DROP TABLE IF EXISTS "+currentTableName+";\n");
        sb.append("CREATE TABLE "+currentTableName);
        sb.append("(\n");
        for(FieldProperty property : fieldProperties) {
            StringBuilder propertyClause = new StringBuilder();
            propertyClause.append(" ${name} ${type}");
            if(!property.isAllowNull()) {
                propertyClause.append(" NOT NULL");
            }
            else {
                propertyClause.append(" NULL");
            }
            if(!stringIsEmpty(property.getDefaultValue())) {
                propertyClause.append(" DEFAULT "+property.getDefaultValue());
            }
            if(!stringIsEmpty(property.getComment())) {
                propertyClause.append(" COMMENT '"+property.getComment()+"'");
            }
            propertyClause.append(",\n");
            String s = propertyClause.toString();
            String propertyName = property.getPropertyName();
            s = s.replace("${name}", propertyName)
                    .replace("${type}", property.getType());
            sb.append(s);
            if(property.isPrimaryKey()) {
                primayKey.append(propertyName +",");
            }
            if(property.isUnique()) {
                uniqueKey.append(" UNIQUE KEY `"+ propertyName +"_UNIQUE` (`"+ propertyName +"`),\n");
            }
        }
        if(primayKey.length() == 0) {
            throw new RuntimeException(currentTableName+"->没有主键");
        }
        else {
            primayKey.deleteCharAt(primayKey.length() - 1);
        }
        sb.append(" PRIMARY KEY ("+primayKey+")");
        if(uniqueKey.length() > 0) {
            sb.append(",\n");
            uniqueKey.deleteCharAt(uniqueKey.length()-2);
            sb.append(uniqueKey);
        }
        else {
            sb.append("\n");
        }
        sb.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='"+tableComment+"';");

        System.out.println(sb.toString());
        try {
            PreparedStatement statement = conn.prepareStatement(sb.toString());
            statement.execute();
        } catch (SQLException e) {
            e.printStackTrace();
        }

    }

    private boolean stringIsEmpty(String str) {
        return str == null || str.equals("");
    }

    public static void main(String[] args) throws SQLException, IOException, ClassNotFoundException {
        String userDirPath = System.getProperty("user.dir");
        String fileSeparator = System.getProperty("file.separator");

        System.out.println("user.dir="+userDirPath);
        Map<String, String> propertyMap = PropertyUtil.getProertyFromFile(userDirPath + fileSeparator + "system.properties");
        String url = propertyMap.get("jdbc.url");
        String account = propertyMap.get("jdbc.account");
        String password = propertyMap.get("jdbc.password");
        String filePath = propertyMap.get("file.path");

        Excel2Db excel2Db = new Excel2Db();
        excel2Db.excel2Db(url, account, password, filePath);
    }
}
