module Pandocxtyle
  module Patches
    refine ::Docx::Elements::Containers::Table do
      def style
        @style ||= @node.xpath("w:tblPr/w:tblStyle/@w:val").first.value
      end

      def style=(new_style)
        @node.xpath("w:tblPr/w:tblStyle/@w:val").first.value = new_style
      end
    end
  end
end
