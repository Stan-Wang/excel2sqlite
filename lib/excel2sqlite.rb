#!/usr/bin/env ruby

require 'roo'
require 'roo-xls'
require 'sqlite3'
require 'thor'
require 'colorize'

module Excel2SQLite
  class Converter
    attr_reader :excel_file, :db_file, :options

    def initialize(excel_file, db_file, options = {})
      @excel_file = excel_file
      @db_file = db_file
      @options = options
      @options[:headers] = true if @options[:headers].nil?
      @column_counter = 0 # 用于生成唯一列名
      @sql_file = options[:sql] ? "#{File.dirname(db_file)}/#{File.basename(db_file, '.*')}.sql" : nil
      @sql_statements = [] if @sql_file
      @mysql_mode = options[:mysql] || false
    end

    def convert
      validate_files
      workbook = open_excel

      if @sql_file
        # 仅生成SQL脚本
        generate_sql_script(workbook)
      else
        # 创建数据库并导入数据
        db = create_database
        workbook.sheets.each do |sheet_name|
          process_sheet(workbook, sheet_name, db)
        end
        puts "转换完成！数据库文件保存在: #{@db_file}".green
      end
    end

    private

    def validate_files
      raise "错误: Excel文件 '#{@excel_file}' 不存在" unless File.exist?(@excel_file)

      if !@options[:force] && (
         (!@sql_file && File.exist?(@db_file)) ||
         (@sql_file && File.exist?(@sql_file))
       )
        file_path = @sql_file || @db_file
        raise "错误: 输出文件 '#{file_path}' 已存在。使用 --force 选项覆盖"
      end
    end

    def open_excel
      puts "打开Excel文件: #{@excel_file}..."
      extension = File.extname(@excel_file).downcase

      case extension
      when '.xlsx', '.xlsm'
        Roo::Excelx.new(@excel_file)
      when '.xls'
        Roo::Excel.new(@excel_file)
      when '.csv'
        Roo::CSV.new(@excel_file)
      when '.ods'
        Roo::OpenOffice.new(@excel_file)
      else
        raise "不支持的文件格式: #{extension}"
      end
    end

    def create_database
      puts "创建SQLite数据库: #{@db_file}..."
      File.delete(@db_file) if File.exist?(@db_file)
      SQLite3::Database.new(@db_file)
    end

    def process_sheet(workbook, sheet_name, db)
      puts "处理工作表: #{sheet_name}..."
      workbook.default_sheet = sheet_name

      # 调试：显示工作表的维度
      puts "工作表维度: #{workbook.first_row}-#{workbook.last_row} 行, #{workbook.first_column}-#{workbook.last_column} 列"

      # 获取表头
      headers = []
      if @options[:headers]
        # 获取原始表头
        raw_headers = workbook.row(1)
        puts "原始表头: #{raw_headers.inspect}"

        # 处理每个表头并保留非空的表头及其索引
        valid_headers = []
        raw_headers.each_with_index do |header, index|
          sanitized = sanitize_column_name(header.to_s)
          next if sanitized.nil? || sanitized.empty?

          valid_headers << [sanitized, index]
        end

        headers = valid_headers
        start_row = 2
      else
        # 如果没有表头，使用列索引作为表头（仍然跳过空列）
        headers = (1..workbook.last_column).map { |i| ["column_#{i}", i - 1] }
                                           .reject { |_, index| workbook.row(1)[index].nil? }
        start_row = 1
      end

      puts "处理后的表头: #{headers.inspect}"

      # 如果没有有效的列，则跳过此工作表
      if headers.empty?
        puts "警告: 工作表 '#{sheet_name}' 没有有效的列，跳过。".yellow
        return
      end

      # 提取列名和索引
      column_names = headers.map { |name, _| name }
      column_indices = headers.map { |_, index| index }

      # 检测数据类型
      column_types = detect_column_types(workbook, column_indices, start_row)

      # 创建表
      table_name = sanitize_table_name(sheet_name)
      create_table(db, table_name, column_names, column_types)

      # 插入数据
      insert_data(db, workbook, table_name, column_names, column_indices, start_row, column_types)
    end

    def sanitize_column_name(name)
      # 将列名转换为有效的SQLite列名
      return '' if name.nil? || name.to_s.strip.empty?

      # 确保每个表头对应唯一的列名
      @used_names ||= {}

      # 清理并获取基础名称
      if name.to_s =~ /\p{Han}/
        # 中文名称，直接使用col_前缀加序号
        @column_counter ||= 0
        @column_counter += 1
        "col_#{@column_counter}"
      else
        # 非中文名称，尝试保留英文
        base = name.to_s.strip.gsub(/\s+/, '_').gsub(/[^a-zA-Z0-9_]/, '').downcase

        # 如果为空或者仅有特殊字符，使用默认名称
        if base.empty?
          @column_counter ||= 0
          @column_counter += 1
          "column_#{@column_counter}"
        elsif @used_names[base]
          # 确保唯一性
          @used_names[base] += 1
          "#{base}_#{@used_names[base]}"
        else
          @used_names[base] = 1
          base
        end
      end
    end

    def sanitize_table_name(name)
      # 将表名转换为有效的SQLite表名
      name.to_s.strip.gsub(/\s+/, '_').gsub(/[^a-zA-Z0-9_]/, '').downcase
    end

    def create_table(db, table_name, headers, column_types = nil)
      columns = if column_types && column_types.length == headers.length
                  headers.zip(column_types).map { |header, type| "\"#{header}\" #{type}" }.join(', ')
                else
                  headers.map { |header| "\"#{header}\" TEXT" }.join(', ')
                end
      db.execute("DROP TABLE IF EXISTS \"#{table_name}\"")
      db.execute("CREATE TABLE \"#{table_name}\" (#{columns})")
    end

    def insert_data(db, workbook, table_name, headers, column_indices, start_row, column_types = nil)
      # 准备插入语句
      placeholders = Array.new(headers.size, '?').join(', ')
      # 修复SQL语句中的引号问题
      columns = headers.map { |h| "\"#{h}\"" }.join(', ')
      insert_sql = "INSERT INTO \"#{table_name}\" (#{columns}) VALUES (#{placeholders})"

      # 添加调试输出
      puts "SQL: #{insert_sql}"
      puts "列名: #{headers.join(', ')}"
      puts "列索引: #{column_indices.join(', ')}"

      # 批量插入数据
      db.transaction

      valid_rows = 0
      empty_rows = 0
      error_rows = 0

      (start_row..workbook.last_row).each do |row_index|
        row_data = workbook.row(row_index)

        # 添加更多调试信息
        puts "原始行数据 (#{row_index}): #{row_data.inspect}" if valid_rows < 2

        # 如果行数据为nil或空数组，跳过
        if row_data.nil? || row_data.empty?
          empty_rows += 1
          next
        end

        # 如果行数据中所有值都为nil或空字符串，跳过
        if row_data.all? { |cell| cell.nil? || (cell.respond_to?(:empty?) && cell.empty?) }
          empty_rows += 1
          next
        end

        # 只获取我们关心的列的数据
        filtered_data = column_indices.map { |i| i < row_data.length ? row_data[i] : nil }

        # 根据列类型转换数据
        filtered_data = if column_types && column_types.length == filtered_data.length
                          filtered_data.zip(column_types).map do |value, type|
                            case type
                            when 'INTEGER', 'INT'
                              if value.is_a?(Numeric)
                                value.to_i
                              else
                                (value.to_s =~ /\A[-+]?\d+\z/ ? value.to_i : value)
                              end
                            when 'REAL', 'DOUBLE'
                              if value.is_a?(Numeric)
                                value.to_f
                              else
                                (value.to_s =~ /\A[-+]?\d*\.\d+\z/ ? value.to_f : value)
                              end
                            when 'DATE'
                              value.is_a?(Date) || value.is_a?(DateTime) || value.is_a?(Time) ? value.to_s : value
                            else
                              case value
                              when Date, DateTime, Time
                                value.to_s
                              else
                                value
                              end
                            end
                          end
                        else
                          # 转换日期或其他特殊类型为字符串
                          filtered_data.map do |value|
                            case value
                            when Date, DateTime, Time
                              value.to_s
                            else
                              value
                            end
                          end
                        end

        # 如果过滤后的数据全部为空，跳过
        if filtered_data.all? { |cell| cell.nil? || (cell.respond_to?(:empty?) && cell.empty?) }
          empty_rows += 1
          next
        end

        # 调试前两行
        puts "过滤后数据 (#{row_index}): #{filtered_data.inspect}" if valid_rows < 2

        # 执行插入
        db.execute(insert_sql, filtered_data)
        valid_rows += 1
      rescue SQLite3::Exception => e
        puts "SQL错误 (行 #{row_index}): #{e.message}".red
        puts "问题数据: #{row_data.inspect if defined?(row_data)}"
        error_rows += 1
      rescue StandardError => e
        puts "处理错误 (行 #{row_index}): #{e.message}".red
        puts "错误类型: #{e.class}"
        puts "问题数据: #{row_data.inspect if defined?(row_data)}"
        error_rows += 1
      end

      db.commit

      puts "已导入 #{valid_rows} 行数据到表 #{table_name}"
      puts "跳过 #{empty_rows} 行空数据"
      puts "处理失败 #{error_rows} 行数据" if error_rows > 0
    end

    def generate_sql_script(workbook)
      @sql_statements = []

      # 添加SQL头部注释
      @sql_statements << '-- Excel2SQLite 生成的SQL脚本'
      @sql_statements << "-- 源文件: #{@excel_file}"
      @sql_statements << "-- 生成时间: #{Time.now}"

      if @mysql_mode
        @sql_statements << '-- MySQL格式'
        @sql_statements << '-- 使用方法: mysql -u 用户名 -p 数据库名 < 此文件'
      else
        @sql_statements << '-- SQLite格式'
        @sql_statements << '-- 使用方法: sqlite3 数据库文件名 < 此文件'
      end

      @sql_statements << ''

      unless @mysql_mode
        @sql_statements << 'BEGIN TRANSACTION;'
        @sql_statements << ''
      end

      # 处理每个工作表
      workbook.sheets.each do |sheet_name|
        process_sheet_for_sql(workbook, sheet_name)
      end

      # 添加提交事务
      unless @mysql_mode
        @sql_statements << ''
        @sql_statements << 'COMMIT;'
      end

      # 写入SQL文件
      File.open(@sql_file, 'w') do |file|
        file.puts @sql_statements.join("\n")
      end

      puts "转换完成！SQL脚本保存在: #{@sql_file}".green
      if @mysql_mode
        puts "可以使用以下命令导入到MySQL: mysql -u 用户名 -p 数据库名 < #{@sql_file}".yellow
      else
        puts "可以使用以下命令导入到SQLite: sqlite3 数据库文件名 < #{@sql_file}".yellow
      end
    end

    def process_sheet_for_sql(workbook, sheet_name)
      puts "处理工作表: #{sheet_name} (SQL生成)..."
      workbook.default_sheet = sheet_name

      # 获取表头
      headers = []
      if @options[:headers]
        raw_headers = workbook.row(1)

        valid_headers = []
        raw_headers.each_with_index do |header, index|
          sanitized = sanitize_column_name(header.to_s)
          next if sanitized.nil? || sanitized.empty?

          valid_headers << [sanitized, index]
        end

        headers = valid_headers
        start_row = 2
      else
        headers = (1..workbook.last_column).map { |i| ["column_#{i}", i - 1] }
                                           .reject { |_, index| workbook.row(1)[index].nil? }
        start_row = 1
      end

      # 如果没有有效的列，则跳过此工作表
      if headers.empty?
        puts "警告: 工作表 '#{sheet_name}' 没有有效的列，跳过。".yellow
        return
      end

      # 提取列名和索引
      column_names = headers.map { |name, _| name }
      column_indices = headers.map { |_, index| index }

      # 预先检查数据类型
      column_types = detect_column_types(workbook, column_indices, start_row)

      # 创建表的SQL语句
      table_name = sanitize_table_name(sheet_name)

      if @mysql_mode
        # MySQL表创建语句
        columns = column_names.zip(column_types).map { |header, type| "`#{header}` #{type}" }.join(', ')
        @sql_statements << "-- 表: #{table_name}"
        @sql_statements << "DROP TABLE IF EXISTS `#{table_name}`;"
        @sql_statements << "CREATE TABLE `#{table_name}` (#{columns}) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;"
      else
        # SQLite表创建语句
        columns = column_names.zip(column_types).map { |header, type| "\"#{header}\" #{type}" }.join(', ')
        @sql_statements << "-- 表: #{table_name}"
        @sql_statements << "DROP TABLE IF EXISTS \"#{table_name}\";"
        @sql_statements << "CREATE TABLE \"#{table_name}\" (#{columns});"
      end

      @sql_statements << ''

      # 生成INSERT语句
      valid_rows = 0
      empty_rows = 0

      (start_row..workbook.last_row).each do |row_index|
        row_data = workbook.row(row_index)

        # 跳过空行
        next if row_data.nil? || row_data.empty?
        next if row_data.all? { |cell| cell.nil? || (cell.respond_to?(:empty?) && cell.empty?) }

        # 只获取我们关心的列的数据
        filtered_data = column_indices.map { |i| i < row_data.length ? row_data[i] : nil }

        # 转换日期或其他特殊类型为字符串
        filtered_data = filtered_data.map do |value|
          case value
          when Date, DateTime, Time
            value.to_s
          else
            value
          end
        end

        # 跳过全空的数据行
        next if filtered_data.all? { |cell| cell.nil? || (cell.respond_to?(:empty?) && cell.empty?) }

        # 格式化SQL中的值（处理NULL、字符串转义等）
        sql_values = filtered_data.zip(column_types).map do |val, type|
          format_sql_value(val, type)
        end

        # 生成INSERT语句
        @sql_statements << if @mysql_mode
                             "INSERT INTO `#{table_name}` (#{column_names.map do |c|
                               "`#{c}`"
                             end.join(', ')}) VALUES (#{sql_values.join(', ')});"
                           else
                             "INSERT INTO \"#{table_name}\" (#{column_names.map do |c|
                               "\"#{c}\""
                             end.join(', ')}) VALUES (#{sql_values.join(', ')});"
                           end
        valid_rows += 1
      rescue StandardError => e
        puts "处理错误 (行 #{row_index}): #{e.message}".red
        empty_rows += 1
      end

      @sql_statements << ''
      puts "已生成 #{valid_rows} 行INSERT语句"
      puts "跳过 #{empty_rows} 行空数据或无效数据"
    end

    # 检测列的数据类型
    def detect_column_types(workbook, column_indices, start_row)
      sample_size = [30, workbook.last_row - start_row + 1].min
      column_types = Array.new(column_indices.size, 'TEXT')

      column_indices.each_with_index do |col_idx, idx|
        integer_count = 0
        float_count = 0
        date_count = 0
        null_count = 0

        # 采样检查数据类型
        (start_row...[start_row + sample_size, workbook.last_row + 1].min).each do |row_idx|
          value = row_idx <= workbook.last_row ? workbook.row(row_idx)[col_idx] : nil

          if value.nil? || (value.respond_to?(:empty?) && value.empty?)
            null_count += 1
            next
          end

          case value
          when Integer
            integer_count += 1
          when Float
            float_count += 1
          when Date, DateTime, Time
            date_count += 1
          when String
            # 尝试转换为数字
            if value =~ /\A[-+]?\d+\z/
              integer_count += 1
            elsif value =~ /\A[-+]?\d*\.\d+\z/
              float_count += 1
            # 增强日期识别能力
            elsif value =~ /\A\d{4}-\d{2}-\d{2}\z/ || # YYYY-MM-DD
                  value =~ %r{\A\d{4}/\d{2}/\d{2}\z} ||                       # YYYY/MM/DD
                  value =~ %r{\A\d{2}/\d{2}/\d{4}\z} ||                       # MM/DD/YYYY
                  value =~ /\A\d{2}-\d{2}-\d{4}\z/ || # MM-DD-YYYY
                  value =~ /\A\d{4}年\d{1,2}月\d{1,2}日\z/ || # 中文日期格式
                  value =~ /\A\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\z/ || # YYYY-MM-DD HH:MM:SS
                  value =~ /\A\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:Z|[-+]\d{2}:?\d{2})?\z/ # ISO 8601
              date_count += 1
            end
          end
        end

        # 确定数据类型，需要一定比例的样本符合才能判定
        non_null_samples = sample_size - null_count
        next unless non_null_samples > 0

        column_types[idx] = if date_count / non_null_samples.to_f >= 0.7
                              @mysql_mode ? 'DATE' : 'DATE'
                            elsif integer_count / non_null_samples.to_f >= 0.7
                              @mysql_mode ? 'INT' : 'INTEGER'
                            elsif (integer_count + float_count) / non_null_samples.to_f >= 0.7
                              @mysql_mode ? 'DOUBLE' : 'REAL'
                            else
                              @mysql_mode ? 'VARCHAR(255)' : 'TEXT'
                            end
      end

      column_types
    end

    # 根据类型将值转换为合适的SQL值
    def format_sql_value(val, type)
      if val.nil?
        'NULL'
      elsif %w[INTEGER INT].include?(type) && val.is_a?(Numeric)
        val.to_i.to_s
      elsif %w[REAL DOUBLE].include?(type) && val.is_a?(Numeric)
        val.to_f.to_s
      elsif type == 'DATE'
        if val.is_a?(Date) || val.is_a?(DateTime) || val.is_a?(Time)
          "'#{val}'"
        elsif val.is_a?(String) && is_date_string?(val)
          "'#{val}'"
        else
          "'#{val.to_s.gsub("'", "''")}'"
        end
      elsif val.is_a?(Numeric)
        val.to_s
      else
        # 转义单引号并包装字符串
        "'#{val.to_s.gsub("'", "''")}'"
      end
    end

    # 检查字符串是否为日期格式
    def is_date_string?(val)
      return false unless val.is_a?(String)

      # 常见日期格式正则匹配
      val =~ /\A\d{4}-\d{2}-\d{2}\z/ || # YYYY-MM-DD
        val =~ %r{\A\d{4}/\d{2}/\d{2}\z} ||                       # YYYY/MM/DD
        val =~ %r{\A\d{2}/\d{2}/\d{4}\z} ||                       # MM/DD/YYYY
        val =~ /\A\d{2}-\d{2}-\d{4}\z/ || # MM-DD-YYYY
        val =~ /\A\d{4}年\d{1,2}月\d{1,2}日\z/ || # 中文日期格式
        val =~ /\A\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\z/ || # YYYY-MM-DD HH:MM:SS
        val =~ /\A\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:Z|[-+]\d{2}:?\d{2})?\z/ # ISO 8601
    end
  end

  class CLI < Thor
    # 默认方法用于处理没有明确子命令的情况
    desc '[EXCEL_FILE] [DB_FILE]', '将Excel文件转换为SQLite数据库'
    method_option :force, type: :boolean, aliases: '-f', desc: '强制覆盖已存在的数据库文件'
    method_option :headers, type: :boolean, default: true, desc: '指定Excel文件是否包含表头'
    method_option :sql, type: :boolean, aliases: '-s', desc: '生成SQL脚本而不是SQLite数据库'
    method_option :mysql, type: :boolean, aliases: '-m', desc: '生成MySQL格式的SQL脚本'
    def default(*args)
      if args.empty?
        help
        return
      end

      excel_file = args[0]
      db_file = args[1]

      # 如果没有指定数据库文件，使用Excel文件的路径和名称（更改扩展名）
      unless db_file
        # 提取源Excel文件的目录和文件名
        excel_dir = File.dirname(excel_file)
        excel_basename = File.basename(excel_file, File.extname(excel_file))
        # 构建输出文件路径
        ext = options[:sql] ? '.sql' : '.db'
        db_file = File.join(excel_dir, "#{excel_basename}#{ext}")
      end

      puts "Excel文件: #{excel_file}"
      puts options[:sql] ? "SQL文件: #{db_file}" : "数据库文件: #{db_file}"
      puts "MySQL格式: #{options[:mysql] ? '是' : '否'}" if options[:sql]

      begin
        # 复制选项，避免修改冻结的Hash
        opts = options.dup
        converter = Converter.new(excel_file, db_file, opts)
        converter.convert
      rescue StandardError => e
        puts "错误: #{e.message}".red
        exit 1
      end
    end

    # 保留convert命令以向后兼容
    desc 'convert EXCEL_FILE [DB_FILE]', '将Excel文件转换为SQLite数据库'
    method_option :force, type: :boolean, aliases: '-f', desc: '强制覆盖已存在的数据库文件'
    method_option :headers, type: :boolean, default: true, desc: '指定Excel文件是否包含表头'
    method_option :sql, type: :boolean, aliases: '-s', desc: '生成SQL脚本而不是SQLite数据库'
    method_option :mysql, type: :boolean, aliases: '-m', desc: '生成MySQL格式的SQL脚本'
    def convert(excel_file, db_file = nil)
      default(excel_file, db_file)
    end

    def self.exit_on_failure?
      true
    end

    # 覆盖help命令，显示更简洁的帮助信息
    def help(command = nil)
      if command.nil?
        puts '用法: excel2sqlite EXCEL_FILE [DB_FILE] [选项]'
        puts ''
        puts '将Excel文件转换为SQLite数据库或SQL脚本'
        puts ''
        puts '选项:'
        puts '  -f, --force                 强制覆盖已存在的数据库或SQL文件'
        puts '  -s, --sql                   生成SQL脚本而不是SQLite数据库'
        puts '  -m, --mysql                 生成MySQL格式的SQL脚本（需与--sql一起使用）'
        puts '  --headers                   指定Excel文件包含表头（默认为true）'
        puts '  --no-headers                指定Excel文件不包含表头'
        puts '  --help                      显示此帮助信息'
        puts ''
        puts '示例:'
        puts '  excel2sqlite data.xlsx              # 创建data.db数据库'
        puts '  excel2sqlite data.xlsx output.db    # 创建指定名称的数据库'
        puts '  excel2sqlite data.xlsx --force      # 覆盖已存在的数据库'
        puts '  excel2sqlite data.xlsx --sql        # 生成SQLite格式的SQL脚本'
        puts '  excel2sqlite data.xlsx --sql --mysql # 生成MySQL格式的SQL脚本'
        puts ''
        puts '也可以使用旧格式(向后兼容):'
        puts '  excel2sqlite convert data.xlsx'
      else
        super
      end
    end
  end
end # 结束 module Excel2SQLite

# 如果直接运行此脚本
if __FILE__ == $0
  # 重新安排参数，将第一个参数作为子命令
  if ARGV.size > 0 && !ARGV[0].start_with?('-') && File.exist?(ARGV[0])
    # 将文件路径作为参数传递给default命令
    ARGV.unshift('default')
  end
  Excel2SQLite::CLI.start(ARGV)
end
