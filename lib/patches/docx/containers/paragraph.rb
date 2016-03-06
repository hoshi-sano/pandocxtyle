module Pandocxtyle
  module Patches
    refine ::Docx::Elements::Containers::Paragraph do
      def properties
        @node.xpath("w:#{@properties_tag}").map do |rpr_node|
          ::Docx::Elements::Containers::ParagraphProperty.new(rpr_node)
        end
      end

      def pstyle_is_a?(pstyle_str)
        properties.any? { |prop| prop.pstyle_value == pstyle_str }
      end

      def append_field_simple(node, attributes = {})
        field_simple = Nokogiri::XML::Node.new("w:fldSimple", @node)
        attributes.each { |key, value| field_simple[key] = value }
        field_simple.add_child(node)
        @node.add_child(field_simple)
        self
      end

      def insert_page_break(pos)
        new_node = Nokogiri::XML::Node.new("w:p", self.parent)
        new_paragraph = self.class.new(new_node)
        text_run = Docx::Elements::Containers::TextRun.create_with(new_paragraph)
        br = Nokogiri::XML::Node.new("w:br", text_run.node)
        br["w:type"] = "page"
        text_run.node.add_child(br)
        text_run.append_to(new_paragraph)
        case pos
        when :before
          new_paragraph.insert_before(self)
        when :after
          new_paragraph.insert_after(self)
        end
      end
    end
  end
end
