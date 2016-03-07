module Pandocxtyle
  using Patches

  module Document
    module_function
    def add_index
      params = ARGV.getopts("i:o:")
      document_properties = %w(Title Author Date)
      Docx::Document.open(params["i"]) do |doc|
        first_content = doc.paragraphs.find do |paragraph|
          document_properties.none? do |property_name|
            paragraph.pstyle_is_a?(property_name)
          end
        end
        new_paragraph = Docx::Elements::Containers::Paragraph.new(
          Nokogiri::XML::Node.new("w:p", first_content.parent)
        )
        text_run = Docx::Elements::Containers::TextRun.create_with(new_paragraph)
        text_run.append_fld_char("begin", {"w:dirty" => "true"})
        text_run.append_fld_char("separate")
        text_run.append_fld_char("end")
        new_paragraph.append_field_simple(text_run.node,
                                          {"w:instr" => " TOC \\o \"1-3\" \\u "})
        new_paragraph.insert_before(first_content)
        new_paragraph.insert_page_break(:before)
        doc.save(params["o"])
      end
    end

    def pagebreak_h1
      params = ARGV.getopts("i:o:")
      Docx::Document.open(params["i"]) do |doc|
        doc.paragraphs.each do |paragraph|
          next unless paragraph.pstyle_is_a?("Heading1")
          paragraph.insert_page_break(:before)
        end
        doc.save(params["o"])
      end
    end
  end
end
