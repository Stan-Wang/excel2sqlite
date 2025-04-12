#!/usr/bin/env ruby

require_relative 'lib/excel2sqlite'
require 'fileutils'

# 测试MySQL格式SQL导出功能

# 如果没有examples目录，创建一个
FileUtils.mkdir_p('examples') unless Dir.exist?('examples')

# 测试文件路径
sample_excel = 'examples/sample.xlsx'
output_sql = 'examples/sample_mysql.sql'

# 清理旧文件
FileUtils.rm(output_sql) if File.exist?(output_sql)

# 判断测试Excel文件是否存在
if File.exist?(sample_excel)
  puts "使用已存在的测试文件: #{sample_excel}"
else
  puts "测试文件不存在，请先创建测试数据的Excel文件: #{sample_excel}"
  puts '您可以运行 ruby examples/create_sample.rb 创建示例文件'
  exit 1
end

puts '=== 测试MySQL格式SQL导出 ==='
begin
  options = {
    force: true,
    headers: true,
    sql: true,
    mysql: true
  }

  converter = Excel2SQLite::Converter.new(sample_excel, output_sql, options.dup)
  converter.convert
  puts '✓ MySQL格式SQL转换成功'

  # 显示生成的SQL脚本的前几行
  if File.exist?(output_sql)
    puts "\n=== MySQL SQL脚本预览 ==="
    first_lines = File.readlines(output_sql)[0..20]
    puts first_lines.join('')
    puts '...'

    puts "\n=== MySQL表结构预览 ==="
    create_table_lines = File.readlines(output_sql).select { |line| line.include?('CREATE TABLE') }
    puts create_table_lines.join("\n")

    puts "\n=== MySQL INSERT预览 ==="
    insert_lines = File.readlines(output_sql).select { |line| line.include?('INSERT INTO') }.first(3)
    puts insert_lines.join("\n")
  end
rescue StandardError => e
  puts "✗ MySQL格式SQL转换失败: #{e.message}"
  puts e.backtrace
end

puts "\n现在可以使用以下命令将数据导入到MySQL数据库："
puts "  mysql -u 用户名 -p 数据库名 < #{output_sql}"
puts "\n测试完成！"
