lib = File.expand_path('lib', __dir__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

Gem::Specification.new do |spec|
  spec.name = 'excel2sqlite'
  spec.version       = '0.2.1'
  spec.authors       = ['Stan Wang']
  spec.email         = ['stanwanng@gmail.com']
  spec.summary       = '将Excel文件转换为SQLite或MySQL数据库的工具'
  spec.description   = '一个Ruby工具，用于将Excel文件(.xlsx, .xls等)转换为SQLite数据库文件或SQL脚本(SQLite/MySQL格式)'
  spec.homepage      = 'https://github.com/Stan-Wang/excel2sqlite'
  spec.license       = 'MIT'

  spec.files         = Dir.glob('{bin,lib}/**/*') + %w[README.md]
  spec.executables   = ['excel2sqlite']
  spec.require_paths = ['lib']

  spec.add_dependency 'colorize', '~> 0.8.1'
  spec.add_dependency 'roo', '~> 2.9.0'
  spec.add_dependency 'roo-xls', '~> 1.2.0'
  spec.add_dependency 'sqlite3', '~> 1.6.0'
  spec.add_dependency 'thor', '~> 1.2.1'

  spec.add_development_dependency 'bundler', '~> 2.0'
  spec.add_development_dependency 'rake', '~> 13.0'
end
