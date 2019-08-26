# Excel表格导出数据库

在excel中按照一定的格式描述表信息，然后按照excel中描述的内容把表的格式导入数据库中。

# Excel格式

参考/src/main/resource/2008.xlsx文件

![参考图片][excelformat.png]

# 配置文件
在文件根目下找到system.properties配置文件

配置内容包括：

```
jdbc.url=jdbc:mysql://ip:port/dbname?useUnicode=true&characterEncoding=utf-8&allowMultiQueries=true&serverTimezone=UTC&useSSL=false
jdbc.account=account
jdbc.password=password
file.path=D:/2008.xlsx
```

- **jdbc.url** 配置数据库连接字符串
- **jdbc.account** 配置访问数据库用户
- **jdbc.password** 配置数据库访问用户的密码
- **file.path** excel配置文件位置（绝对路径）

# 打包

- 在根目录下，执行mvn clean package -DskipTests

- 在target目录下生成了excel2db-1.0-SNAPSHOT.jar文件

# 运行

1. 将excel2db-1.0-SNAPSHOT.jar文件和system.properties拷贝到同一目录下

2. 在命令行中执行java -jar excel2db-1.0-SNAPSHOT.jar即可启动
