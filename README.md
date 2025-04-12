# Excel2SQLite 转换工具

一个简单易用的Ruby工具，用于将Excel文件转换为SQLite数据库或SQL脚本（支持SQLite和MySQL格式）。

## 功能特点

- 支持多种Excel格式（.xlsx, .xls, .csv, .ods）
- 自动识别表头
- 将Excel中的每个工作表转换为SQLite中的表
- 简单的命令行界面
- 可作为系统命令使用
- 自动处理中文表头和特殊字符
- 跳过空行和空列
- 支持日期类型数据转换
- 可导出SQL脚本文件而非直接创建数据库
- 智能数据类型检测（INTEGER、REAL、DATE、TEXT）
- 支持MySQL格式的SQL脚本导出

## 安装

### 方法1：从RubyGems安装

```bash
gem install excel2sqlite
```

### 方法2：从源代码安装

```bash
git clone https://github.com/Stan-Wang/excel2sqlite.git
cd excel2sqlite
bundle install
rake install
```

## 使用方法

基本用法：

```bash
excel2sqlite path/to/your/file.xlsx [output.db]
```

如果不指定输出文件名，将使用Excel文件名（更改扩展名为.db或.sql）作为输出文件名，且生成在Excel文件的同一目录下。

### 选项

- `-f, --force`：强制覆盖已存在的数据库或SQL文件
- `-s, --sql`：生成SQL脚本而不是SQLite数据库
- `-m, --mysql`：生成MySQL格式的SQL脚本（需与`--sql`一起使用）
- `--no-headers`：指定Excel文件不包含表头（默认为true）
- `--help`：显示帮助信息

### 示例

```bash
# 基本转换（生成数据库）
excel2sqlite data.xlsx

# 指定输出文件名
excel2sqlite data.xlsx output.db

# 强制覆盖已存在的文件
excel2sqlite data.xlsx --force

# 指定Excel文件不包含表头
excel2sqlite data.xlsx --no-headers

# 生成SQLite格式的SQL脚本
excel2sqlite data.xlsx --sql

# 生成MySQL格式的SQL脚本
excel2sqlite data.xlsx --sql --mysql

# 生成指定名称的MySQL格式SQL脚本
excel2sqlite data.xlsx output.sql --sql --mysql
```

## 数据处理说明

- 工具会自动跳过空的行和列
- 中文表头会被转换为有效的SQLite表头（格式为col_数字）
- 非ASCII字符会被移除或替换为下划线
- 每个工作表会转换为一个SQLite表，表名为工作表名的小写版本（空格会被替换为下划线）
- 日期类型数据会被转换为字符串格式
- 工具通过采样分析自动检测列的数据类型：
  - SQLite模式:
    - INTEGER：整数类型数据
    - REAL：浮点数类型数据
    - DATE：日期类型数据
    - TEXT：文本或其他类型数据
  - MySQL模式:
    - INT：整数类型数据
    - DOUBLE：浮点数类型数据
    - DATE：日期类型数据
    - VARCHAR(255)：文本或其他类型数据

## SQL脚本模式

使用`--sql`选项时，工具会生成包含CREATE TABLE和INSERT语句的SQL脚本文件，而不是直接创建数据库。这样您可以：

- 手动检查和修改SQL语句
- 在不同的数据库实例中重复使用
- SQLite模式：将脚本导入到数据库：`sqlite3 数据库文件名 < 脚本.sql`
- MySQL模式：将脚本导入到数据库：`mysql -u 用户名 -p 数据库名 < 脚本.sql`

### MySQL模式

当使用`--sql --mysql`选项时，工具会生成适用于MySQL的SQL脚本：

- 使用MySQL兼容的语法
- 使用反引号(`)而不是双引号来包围表名和列名
- 将数据类型映射到MySQL等效类型
- 添加MySQL特定的表属性（如字符集、引擎）
- 不添加BEGIN TRANSACTION和COMMIT语句

## 开发

### 依赖项

- Ruby >= 2.5.0
- roo
- roo-xls
- sqlite3
- thor
- colorize

### 测试

```bash
rake test
```

## 许可证

MIT