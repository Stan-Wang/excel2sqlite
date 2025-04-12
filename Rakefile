require 'bundler/gem_tasks'

task :default => :build

desc "安装开发依赖"
task :setup do
  sh "bundle install"
end

desc "本地安装gem"
task :install_local => :build do
  sh "gem install --local pkg/excel2sqlite-0.1.0.gem"
end

desc "创建示例Excel文件用于测试"
task :create_sample do
  require 'axlsx'
  
  p = Axlsx::Package.new
  wb = p.workbook
  
  # 创建第一个工作表
  wb.add_worksheet(name: '员工') do |sheet|
    sheet.add_row ['ID', '姓名', '部门', '薪资']
    sheet.add_row [1, '张三', '技术部', 10000]
    sheet.add_row [2, '李四', '市场部', 8500]
    sheet.add_row [3, '王五', '人事部', 7800]
  end
  
  # 创建第二个工作表
  wb.add_worksheet(name: '部门') do |sheet|
    sheet.add_row ['部门ID', '部门名称', '负责人']
    sheet.add_row [1, '技术部', '张总']
    sheet.add_row [2, '市场部', '李总']
    sheet.add_row [3, '人事部', '王总']
  end
  
  p.serialize('sample.xlsx')
  puts "已创建示例Excel文件: sample.xlsx"
end