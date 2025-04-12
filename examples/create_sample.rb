#!/usr/bin/env ruby

require 'axlsx'

# 创建一个示例Excel文件
p = Axlsx::Package.new
wb = p.workbook

# 创建第一个工作表 - 员工信息
wb.add_worksheet(name: '员工') do |sheet|
  sheet.add_row %W[ID \u59D3\u540D \u90E8\u95E8 \u85AA\u8D44]
  sheet.add_row [1, '张三', '技术部', 10_000]
  sheet.add_row [2, '李四', '市场部', 8500]
  sheet.add_row [3, '王五', '人事部', 7800]
  sheet.add_row [4, '赵六', '技术部', 12_000]
  sheet.add_row [5, '钱七', '市场部', 9000]
end

# 创建第二个工作表 - 部门信息
wb.add_worksheet(name: '部门') do |sheet|
  sheet.add_row %W[\u90E8\u95E8ID \u90E8\u95E8\u540D\u79F0 \u8D1F\u8D23\u4EBA \u6210\u7ACB\u65E5\u671F]
  sheet.add_row [1, '技术部', '张总', '2020-01-15']
  sheet.add_row [2, '市场部', '李总', '2020-02-20']
  sheet.add_row [3, '人事部', '王总', '2020-03-10']
end

# 保存文件
output_file = File.join(File.dirname(__FILE__), 'sample.xlsx')
p.serialize(output_file)
puts "已创建示例Excel文件: #{output_file}"
