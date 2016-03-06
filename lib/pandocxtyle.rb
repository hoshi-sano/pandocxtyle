require "optparse"
require "docx"
require "docx/document"
require "patches"
require "pandocxtyle/version"
require "pandocxtyle/document"
require "pandocxtyle/table"
require "pandocxtyle/drawing"

module Pandocxtyle
  COMMAND_WHITELIST = %w(document table drawing)
  ACTION_WHITELIST = {
    "document" => %w(add_index pagebreak_h1),
    "table"    => %w(sub_style seq_caption),
    "drawing"  => %w(seq_caption),
  }
  EXTRA_ACTIONS = %w(-v --version -h --help)
  HELP_MESSAGE = <<-EOS
  USAGE:
    pandocxtyle [options]
    pandocxtyle <command> <action> [options]

  GETTING HELP / INFORMATION:
    -h, --help                show help
    -v, --version             show version

  COMMANDS:
    document
    table
    drawing

  GLOBAL OPTIONS:
    -i                        input docx file path
    -o                        output docx file path

  ACTIONS AND OPTIONS:
    document
      add_index               insert index before first paragraph
      pagebreak_h1            insert page break before Heading1

    table
      sub_style               substitute the table style
        -t, --target          target style id for substitution
                              DEFAULT: '#{Table::DEFAULT_TARGET_STYLE}'
        -w, --with            new style id
      seq_caption             add sequential number field to table captions
        -l, --label           DEFAULT: '#{Table::DEFAULT_CAPTION_LABEL}'

    drawing
      seq_caption             add sequential number field to image captions
        -l, --label           DEFAULT: '#{Drawing::DEFAULT_CAPTION_LABEL}'
  EOS

  module_function

  def execute
    write_help_or_version if (EXTRA_ACTIONS & ARGV).any?
    debug = ARGV.delete("--debug")

    command = ARGV.shift
    unless COMMAND_WHITELIST.include?(command)
      write_error_message(command)
    end
    action = ARGV.shift
    unless ACTION_WHITELIST[command].include?(action)
      write_error_message(command, action)
    end
    const_get(camelize(command)).send(action)
  rescue => e
    puts e.message
    puts e.backtrace.join("\n") if debug
  end

  def write_help_or_version
    params = ARGV.getopts("hv", "help", "version")
    if params["v"] || params["version"]
      puts VERSION
    end
    if params["h"] || params["help"]
      puts HELP_MESSAGE
    end
    exit(0)
  end

  def write_error_message(command, action = nil)
    if action
      puts "No such action '#{action}' for #{command}"
    else
      puts "No such command '#{command}'"
    end
    puts "See 'pandocxtyle --help' for information."
    exit(1)
  end

  def camelize(str)
    str.split(/[^[:alnum:]]+/).map(&:capitalize).join
  end
end
