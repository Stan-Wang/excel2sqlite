#!/usr/bin/env ruby

require_relative 'lib/excel2sqlite'
require 'fileutils'

# 测试数据类型检测功能

# 如果没有examples目录，创建一个
FileUtils.mkdir_p('examples') unless Dir.exist?('examples')

# 测试文件路径
test_excel = 'examples/test_data_types.xlsx'
test_db = 'examples/test_data_types.db'
test_sql = 'examples/test_data_types.sql'

# 清理旧文件
FileUtils.rm(test_db) if File.exist?(test_db)
FileUtils.rm(test_sql) if File.exist?(test_sql)

# 判断测试Excel文件是否存在
if File.exist?(test_excel)
  puts "测试文件已存在: #{test_excel}"
else
  puts "测试文件不存在，请先创建测试数据的Excel文件: #{test_excel}"
  puts '测试文件应包含不同类型的数据列：整数、浮点数、日期和文本'
  exit 1
end

puts '=== 测试1: 转换为数据库 ==='
begin
  options = { force: true, headers: true }
  converter = Excel2SQLite::Converter.new(test_excel, test_db, options.dup)
  converter.convert
  puts '✓ 转换为数据库成功'
rescue StandardError => e
  puts "✗ 转换为数据库失败: #{e.message}"
  puts e.backtrace
end

puts "\n=== 测试2: 转换为SQL脚本 ==="
begin
  options = { force: true, headers: true, sql: true }
  converter = Excel2SQLite::Converter.new(test_excel, test_sql, options.dup)
  converter.convert
  puts '✓ 转换为SQL脚本成功'

  # 显示生成的SQL脚本的前几行
  if File.exist?(test_sql)
    puts "\n=== SQL脚本预览 ==="
    first_lines = File.readlines(test_sql)[0..20]
    puts first_lines.join('')
    puts '...'
  end
rescue StandardError => e
  puts "✗ 转换为SQL脚本失败: #{e.message}"
  puts e.backtrace
end

puts "\n测试完成！"
