#!/bin/bash

echo "===== 安装 Excel2SQLite 工具 ====="

# 安装依赖
echo "正在安装依赖..."
gem install colorize -v "~> 0.8.1" 2>/dev/null || sudo gem install colorize -v "~> 0.8.1"
gem install thor -v "~> 1.2.1" 2>/dev/null || sudo gem install thor -v "~> 1.2.1"
gem install roo -v "~> 2.9.0" 2>/dev/null || sudo gem install roo -v "~> 2.9.0"
gem install roo-xls -v "~> 1.2.0" 2>/dev/null || sudo gem install roo-xls -v "~> 1.2.0"
gem install sqlite3 -v "~> 1.6.0" 2>/dev/null || sudo gem install sqlite3 -v "~> 1.6.0"

# 创建pkg目录（如果不存在）
mkdir -p pkg

# 构建gem包
echo "正在构建gem包..."
gem build -o pkg/excel2sqlite-0.1.0.gem excel2sqlite.gemspec 2>/dev/null || {
  echo "构建gem失败，尝试使用系统Ruby..."
  /usr/bin/ruby -S gem build -o pkg/excel2sqlite-0.1.0.gem excel2sqlite.gemspec
}

# 安装gem
echo "正在安装gem..."
gem install --local pkg/excel2sqlite-0.1.0.gem 2>/dev/null || {
  echo "安装gem失败，尝试使用系统Ruby和sudo..."
  sudo /usr/bin/ruby -S gem install --local pkg/excel2sqlite-0.1.0.gem --ignore-dependencies
}

echo "安装完成！"
echo "现在你可以使用 'excel2sqlite convert 文件路径.xlsx' 来转换Excel文件" 