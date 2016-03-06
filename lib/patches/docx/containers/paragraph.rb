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
    end
  end
end
