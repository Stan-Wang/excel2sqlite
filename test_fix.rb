#!/usr/bin/env ruby

require 'thor'

class MockConverter
  def initialize(options)
    @options = options
    @options[:headers] = true if @options[:headers].nil?
  end

  def convert
    puts "选项: #{@options.inspect}"
    puts '成功转换！'
  end
end

class TestCLI < Thor
  desc 'test', '测试冻结的Hash修复'
  method_option :force, type: :boolean, aliases: '-f', desc: '测试选项'

  def test
    # 复制选项，避免修改冻结的Hash
    opts = options.dup
    converter = MockConverter.new(opts)
    converter.convert
    puts '无错误，修复成功！'
  rescue StandardError => e
    puts "错误: #{e.message}"
    exit 1
  end
end

TestCLI.start(ARGV)
